"""FLauncher — F-Engineering Launcher v3 как окно Windows.

Без консоли: проверяет и поднимает backend-сервер в фоне (порты 8780..8789)
и открывает интерфейс в отдельном окне приложения через pywebview.
При отсутствии pywebview или ошибке вебвью открывает окно в Edge/Chrome.
Гарантированное завершение: использует Windows Job Object (KILL_ON_JOB_CLOSE)
и явные хуки закрытия окна, чтобы сервер никогда не оставался зомби в памяти.
"""
import atexit
import ctypes
from ctypes import wintypes
import json
import os
import shutil
import socket
import subprocess
import sys
import time
import urllib.request
from pathlib import Path

if getattr(sys, 'frozen', False):
    # Запущено из скомпилированного EXE (dist\FEngineeringLauncher\FEngineeringLauncher.exe)
    # Корнем лаунчера является папка с самим EXE или папка на уровень выше
    REPO_ROOT = str(Path(sys.executable).resolve().parent)
else:
    # Запущено обычным python app/flauncher.pyw
    REPO_ROOT = str(Path(__file__).resolve().parent.parent)


def _resource_path(*relative: str) -> str:
    """Найти bundled-ресурс: сначала от REPO_ROOT, затем из sys._MEIPASS.

    PyInstaller 6 в режиме onedir кладёт --add-data в подпапку _internal,
    на которую указывает sys._MEIPASS. В обычном режиме _MEIPASS нет
    и всё строго от REPO_ROOT.
    """
    candidate = os.path.join(REPO_ROOT, *relative)
    if os.path.exists(candidate):
        return candidate
    meipass = getattr(sys, '_MEIPASS', '')
    if meipass:
        bundled = os.path.join(str(meipass), *relative)
        if os.path.exists(bundled):
            return bundled
    return candidate


def _backend_python() -> str:
    """Интерпретатор для запуска backend-сервера (server.py).

    Приоритет: встроенный runtime/python/pythonw.exe, затем (только
    в frozen-режиме без embedded runtime) pythonw/python из PATH.
    В обычном режиме — текущий интерпретатор. Никогда не возвращаем
    сам EXE: иначе frozen-процесс породил бы каскад копий GUI.
    """
    if os.path.exists(EMBEDDED_PYTHON):
        return EMBEDDED_PYTHON
    if getattr(sys, 'frozen', False):
        for candidate in ("pythonw.exe", "python.exe"):
            found = shutil.which(candidate)
            if found:
                return found
    return sys.executable


BACKEND = _resource_path("app", "backend", "server.py")
ICON = _resource_path("app", "frontend", "assets", "flauncher.ico")
EMBEDDED_PYTHON = os.path.join(REPO_ROOT, "runtime", "python", "pythonw.exe")
TITLE = "F-Engineering Launcher"
GUI_MUTEX_PORT = 8799
LOGS_DIR = os.path.join(REPO_ROOT, "runtime", "logs")

# --- Windows Job Object (Process Lifecycle Management) ---
JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE = 0x2000
JobObjectExtendedLimitInformation = 9


class IO_COUNTERS(ctypes.Structure):
    _fields_ = [
        ("ReadOperationCount", ctypes.c_uint64),
        ("WriteOperationCount", ctypes.c_uint64),
        ("OtherOperationCount", ctypes.c_uint64),
        ("ReadTransferCount", ctypes.c_uint64),
        ("WriteTransferCount", ctypes.c_uint64),
        ("OtherTransferCount", ctypes.c_uint64),
    ]


class JOBOBJECT_BASIC_LIMIT_INFORMATION(ctypes.Structure):
    _fields_ = [
        ("PerProcessUserTimeLimit", ctypes.c_int64),
        ("PerJobUserTimeLimit", ctypes.c_int64),
        ("LimitFlags", wintypes.DWORD),
        ("MinimumWorkingSetSize", ctypes.c_size_t),
        ("MaximumWorkingSetSize", ctypes.c_size_t),
        ("ActiveProcessLimit", wintypes.DWORD),
        ("Affinity", ctypes.c_size_t),
        ("PriorityClass", wintypes.DWORD),
        ("SchedulingClass", wintypes.DWORD),
    ]


class JOBOBJECT_EXTENDED_LIMIT_INFORMATION(ctypes.Structure):
    _fields_ = [
        ("BasicLimitInformation", JOBOBJECT_BASIC_LIMIT_INFORMATION),
        ("IoInfo", IO_COUNTERS),
        ("ProcessMemoryLimit", ctypes.c_size_t),
        ("JobMemoryLimit", ctypes.c_size_t),
        ("PeakProcessMemoryLimit", ctypes.c_size_t),
        ("PeakJobMemoryLimit", ctypes.c_size_t),
    ]


_JOB_HANDLE = None


def get_or_create_kill_on_close_job():
    global _JOB_HANDLE
    if _JOB_HANDLE is not None:
        return _JOB_HANDLE
    if os.name != "nt":
        return None
    try:
        k32 = ctypes.windll.kernel32
        job = k32.CreateJobObjectW(None, None)
        if not job:
            log("CreateJobObjectW returned NULL")
            return None
        info = JOBOBJECT_EXTENDED_LIMIT_INFORMATION()
        info.BasicLimitInformation.LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        success = k32.SetInformationJobObject(
            job,
            JobObjectExtendedLimitInformation,
            ctypes.byref(info),
            ctypes.sizeof(info),
        )
        if not success:
            err = k32.GetLastError()
            log(f"SetInformationJobObject failed with error {err}")
            k32.CloseHandle(job)
            return None
        _JOB_HANDLE = job
        log("Successfully created Windows Job Object with KILL_ON_JOB_CLOSE")
        return _JOB_HANDLE
    except Exception as e:
        log(f"Failed to create Job Object: {e}")
        return None


def assign_process_to_job(proc) -> bool:
    if not proc or os.name != "nt":
        return False
    job = get_or_create_kill_on_close_job()
    if not job:
        return False
    try:
        k32 = ctypes.windll.kernel32
        proc_handle = getattr(proc, "_handle", None)
        if proc_handle:
            ok = bool(k32.AssignProcessToJobObject(job, int(proc_handle)))
            log(f"Assigned process pid {proc.pid} to Job Object: {ok}")
            return ok
        elif hasattr(proc, "pid"):
            PROCESS_SET_QUOTA = 0x0100
            PROCESS_TERMINATE = 0x0001
            h = k32.OpenProcess(PROCESS_SET_QUOTA | PROCESS_TERMINATE, False, proc.pid)
            if h:
                ok = bool(k32.AssignProcessToJobObject(job, h))
                k32.CloseHandle(h)
                log(f"Assigned process pid {proc.pid} to Job Object via OpenProcess: {ok}")
                return ok
    except Exception as e:
        log(f"AssignProcessToJobObject error: {e}")
    return False


def hidden_process_kwargs() -> dict:
    """Флаги скрытого запуска процессов: никаких чёрных окон консоли."""
    kwargs = {}
    if os.name == "nt":
        kwargs["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        si.wShowWindow = 0  # SW_HIDE
        kwargs["startupinfo"] = si
    return kwargs


def kill_process_tree_silent(pid: int) -> None:
    """Завершить процесс и его потомков абсолютно тихо, без консольных утилит."""
    if not pid:
        return
    # 1. Попытка через psutil (быстро, нативно, в памяти процесса)
    try:
        import psutil
        parent = psutil.Process(pid)
        for child in parent.children(recursive=True):
            try:
                child.kill()
            except Exception:
                pass
        parent.kill()
        return
    except Exception:
        pass

    # 2. Попытка через Win32 TerminateProcess API (без внешних процессов и окон)
    if os.name == "nt":
        try:
            k32 = ctypes.windll.kernel32
            handle = k32.OpenProcess(0x0001, False, int(pid))  # PROCESS_TERMINATE = 1
            if handle:
                k32.TerminateProcess(handle, 1)
                k32.CloseHandle(handle)
                return
        except Exception:
            pass

    # 3. Резерв: taskkill со скрытием окна
    if os.name == "nt":
        try:
            subprocess.run(
                ["taskkill", "/F", "/T", "/PID", str(pid)],
                capture_output=True,
                timeout=5,
                **hidden_process_kwargs(),
            )
        except Exception:
            pass


def get_pids_on_port(port: int) -> set[int]:
    pids = set()
    try:
        out = subprocess.run(
            ["netstat", "-ano", "-p", "TCP"],
            capture_output=True, text=True, timeout=5,
            **hidden_process_kwargs(),
        ).stdout
        for line in (out or "").splitlines():
            parts = line.split()
            if len(parts) >= 5 and parts[0] == "TCP" \
                    and parts[1].endswith(":%d" % port) and parts[3] == "LISTENING":
                if parts[4].isdigit():
                    pids.add(int(parts[4]))
    except Exception:
        pass
    return pids


def shutdown_server(proc, port: int | None = None) -> None:
    # 1. Завершаем серверный процесс и его дерево процессов без вызова консольных утилит
    if proc:
        try:
            if proc.poll() is None:
                log(f"Shutting down server pid {proc.pid}...")
                kill_process_tree_silent(proc.pid)
                try:
                    proc.wait(timeout=1.0)
                except Exception:
                    pass
                log(f"Server pid {proc.pid} stopped successfully")
        except Exception as e:
            log(f"Error during shutdown_server proc: {e}")

    # 2. Если сервер работал на порту (даже если был найден ранее запущенным)
    if port:
        for pid in get_pids_on_port(port):
            try:
                log(f"Killing listener process pid {pid} on port {port}...")
                kill_process_tree_silent(pid)
            except Exception:
                pass

    # 3. Гарантированно зачищаем любые осиротевшие невидимые фоновые сессии AutoCAD без окон
    if os.name == "nt":
        try:
            subprocess.run(
                ["taskkill", "/F", "/FI", "WINDOWTITLE eq ", "/IM", "acad.exe"],
                capture_output=True,
                timeout=5,
                **hidden_process_kwargs(),
            )
        except Exception:
            pass


def log(msg: str) -> None:
    try:
        os.makedirs(LOGS_DIR, exist_ok=True)
        log_file = os.path.join(LOGS_DIR, "launcher_gui.log")
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {msg}\n")
    except Exception:
        pass


def is_server_healthy(port: int) -> bool:
    try:
        req = urllib.request.Request(f"http://127.0.0.1:{port}/api/health")
        with urllib.request.urlopen(req, timeout=0.8) as resp:
            if resp.status != 200:
                return False
            data = json.loads(resp.read().decode("utf-8"))
            repo = data.get("repoRoot", "")
            if repo and os.path.normcase(os.path.abspath(repo)) != os.path.normcase(os.path.abspath(REPO_ROOT)):
                return False
            return True
    except Exception:
        return False


def is_port_free(port: int) -> bool:
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.bind(("127.0.0.1", port))
            return True
    except OSError:
        return False


def reap_stale_server(port: int) -> None:
    # Порт занят, но сервер на нём не отвечает: гасим зависший процесс,
    # если это точно процесс лаунчера. Чужие процессы не трогаем.
    try:
        out = subprocess.run(
            ["netstat", "-ano", "-p", "TCP"],
            capture_output=True, text=True, timeout=10,
            **hidden_process_kwargs(),
        ).stdout
    except Exception:
        return
    pids = set()
    for line in (out or "").splitlines():
        parts = line.split()
        if len(parts) >= 5 and parts[0] == "TCP" \
                and parts[1].endswith(":%d" % port) and parts[3] == "LISTENING":
            if parts[4].isdigit():
                pids.add(parts[4])
    for pid in pids:
        try:
            kill_process_tree_silent(int(pid))
            log(f"Reaped stale launcher server pid {pid} on port {port}")
        except Exception:
            pass


def find_or_start_server() -> tuple[int, subprocess.Popen | None]:
    # 1. Проверяем, может сервер уже работает на основном порту 8780
    if is_server_healthy(8780):
        log("Found healthy server on primary port 8780")
        return 8780, None

    # Если порт 8780 занят, но НЕ здоров — сносим зомби-процесс
    if not is_port_free(8780):
        log("Port 8780 is occupied but unhealthy; reaping stale server...")
        reap_stale_server(8780)
        time.sleep(0.5)

    # 2. Если 8780 свободен, выбираем его; иначе проверяем диапазон 8781..8789
    target_port = None
    if is_port_free(8780):
        target_port = 8780
    else:
        for p in range(8781, 8790):
            if is_server_healthy(p):
                log(f"Found already healthy server on port {p}")
                return p, None
            if not is_port_free(p):
                reap_stale_server(p)
                time.sleep(0.3)
            if is_port_free(p):
                target_port = p
                break

    if target_port is None:
        target_port = 8780

    log(f"Starting backend on port {target_port}")
    # Запускаем server.py скрыто: никаких чёрных окон консоли
    python_exe = _backend_python()
    server_proc = subprocess.Popen(
        [python_exe, BACKEND, "--port", str(target_port)],
        cwd=REPO_ROOT,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        **hidden_process_kwargs(),
    )

    # Привязываем дочерний процесс сервера к Windows Job Object с авто-завершением
    assign_process_to_job(server_proc)

    deadline = time.time() + 20
    while time.time() < deadline:
        if is_server_healthy(target_port):
            log(f"Server on port {target_port} became healthy (pid {server_proc.pid})")
            return target_port, server_proc
        time.sleep(0.3)

    log(f"Warning: server on port {target_port} health check timed out, attempting to use it anyway")
    return target_port, server_proc


def focus_existing_window() -> bool:
    try:
        user32 = ctypes.windll.user32
        for t in [TITLE, "FLauncher"]:
            hwnd = user32.FindWindowW(None, t)
            if hwnd:
                user32.ShowWindow(hwnd, 9)  # SW_RESTORE
                user32.SetForegroundWindow(hwnd)
                return True
    except Exception:
        pass
    return False


def open_browser_window(url: str) -> None:
    edge_paths = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]
    for p in edge_paths:
        if os.path.exists(p):
            profile = os.path.join(os.environ.get("LOCALAPPDATA", ""), "FEngineeringLauncher", "AppProfile")
            os.makedirs(profile, exist_ok=True)
            subprocess.Popen(
                [p, f"--app={url}", f"--user-data-dir={profile}", "--no-first-run", "--start-maximized"],
                **hidden_process_kwargs(),
            )
            return

    import webbrowser
    webbrowser.open(url)


def main():
    try:
        mutex = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        mutex.bind(("127.0.0.1", GUI_MUTEX_PORT))
    except OSError:
        if focus_existing_window():
            log("Focused existing window, exiting duplicate process")
            return

    port, server_proc = find_or_start_server()
    atexit.register(lambda: shutdown_server(server_proc, port))

    url = f"http://127.0.0.1:{port}/"

    opened_webview = False
    try:
        import webview
        log(f"Opening pywebview window at {url} (maximized)")
        # maximized=True: окно создаётся сразу на весь экран без промежуточного маленького окна
        window = webview.create_window(
            TITLE,
            url,
            maximized=True,
            min_size=(1024, 700),
            background_color='#0f172a',
        )

        # Гарантированное завершение сервера при закрытии окна
        try:
            def _on_window_closed(*args, **kwargs):
                log("Window closed event received, stopping server...")
                shutdown_server(server_proc, port)
            window.events.closed += _on_window_closed
        except Exception as e:
            log(f"Could not bind on_closed event: {e}")

        opened_webview = True
        webview.start(icon=ICON if os.path.exists(ICON) else None)
        # После выхода из webview.start() (окно закрыто)
        shutdown_server(server_proc, port)
    except Exception as e:
        log(f"pywebview failed: {e}, falling back to browser window")
        if not opened_webview:
            open_browser_window(url)


if __name__ == "__main__":
    main()
