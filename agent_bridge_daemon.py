import json
import os
import subprocess
import time

BRIDGE_DIR = os.path.abspath(".agent_bridge")
TASK_FILE = os.path.join(BRIDGE_DIR, "TASK.json")
STATUS_FILE = os.path.join(BRIDGE_DIR, "STATUS.json")

print("[BRIDGE] Слушатель задач активен. Ожидание задач из Git...")

def git_sync():
    try:
        subprocess.run(["git", "pull", "--rebase"], capture_output=True, text=True, timeout=15)
    except Exception as e:
        pass

def git_push_status(msg):
    try:
        subprocess.run(["git", "add", ".agent_bridge/"], timeout=10)
        subprocess.run(["git", "commit", "-m", f"[BRIDGE STATUS] {msg}"], timeout=10)
        subprocess.run(["git", "push"], timeout=20)
        print(f"[BRIDGE] Статус синхронизирован с Git: {msg}")
    except Exception as e:
        print(f"[BRIDGE ERROR] Ошибка push: {e}")

def update_status(status, task_id, logs):
    data = {
        "status": status,
        "current_task_id": task_id,
        "last_updated": time.strftime("%Y-%m-%d %H:%M:%S"),
        "worker": "WINDOWS_LOCAL_AGENT",
        "logs": logs
    }
    with open(STATUS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    git_push_status(f"Task {task_id}: {status}")

# Начальная фиксация в репозитории
if os.path.exists(TASK_FILE):
    git_push_status("Init .agent_bridge protocol")
