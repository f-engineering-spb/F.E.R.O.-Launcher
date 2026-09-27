import os
import sys
import json
import subprocess

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

def pick_exe():
    # Запускаем PowerShell напрямую с окном-владельцем TopMost, чтобы диалог выбора файла гарантированно открывался поверх всех окон
    ps_cmd = (
        "[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false); "
        "$OutputEncoding = [System.Text.UTF8Encoding]::new($false); "
        "Add-Type -AssemblyName System.Windows.Forms; "
        "$dlg = New-Object System.Windows.Forms.OpenFileDialog; "
        '$dlg.Filter = "Исполняемые файлы (*.exe)|*.exe|Все файлы (*.*)|*.*"; '
        '$dlg.Title = "Выберите программу для запуска (.exe)"; '
        "$dlg.RestoreDirectory = $true; "
        "$owner = New-Object System.Windows.Forms.Form; "
        "$owner.TopMost = $true; $owner.ShowInTaskbar = $false; "
        "$null = $owner.Handle; "
        "try { if ($dlg.ShowDialog($owner) -eq [System.Windows.Forms.DialogResult]::OK) { "
        "if ($dlg.FileName -and (Test-Path -LiteralPath $dlg.FileName)) { [Console]::WriteLine($dlg.FileName) } } } "
        "finally { $owner.Dispose(); $dlg.Dispose() }"
    )
    creation_flags = 0
    if os.name == "nt":
        creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
    try:
        res = subprocess.run(
            [
                "powershell.exe",
                "-WindowStyle", "Hidden",
                "-STA",
                "-NoProfile",
                "-ExecutionPolicy", "Bypass",
                "-Command", ps_cmd
            ],
            capture_output=True,
            text=True,
            encoding="utf-8",
            timeout=120,
            creationflags=creation_flags
        )
        if res.returncode == 0:
            lines = [line.strip() for line in res.stdout.splitlines() if line.strip()]
            valid = [p for p in lines if os.path.isfile(p)]
            if valid:
                return valid[0]
    except Exception:
        pass

    # Фолбэк Tkinter
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        selected = filedialog.askopenfilename(
            title="Выберите программу для запуска (.exe)",
            filetypes=[("Исполняемые файлы", "*.exe"), ("Все файлы", "*.*")]
        )
        root.destroy()
        if selected and os.path.isfile(selected):
            return selected
    except Exception:
        pass
    return ""

if __name__ == "__main__":
    exe = pick_exe()
    if exe:
        print(json.dumps(exe, ensure_ascii=False))
