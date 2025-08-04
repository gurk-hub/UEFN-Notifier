import os
import sys
import time
import winsound
import threading
import json
import glob
import subprocess
from datetime import datetime
from tkinter import Tk, filedialog
import pystray
from pystray import MenuItem as item
from PIL import Image

# For creating Windows shortcuts
import pythoncom
from win32com.shell import shell

# For notifications
from winotify import Notification, audio

# ---------------- PATHS ----------------
APPDATA_FOLDER = os.path.join(os.getenv("APPDATA"), "UEFNPushNotifier")
os.makedirs(APPDATA_FOLDER, exist_ok=True)
SETTINGS_FILE = os.path.join(APPDATA_FOLDER, "settings.json")
EVENT_LOG_FILE = os.path.join(APPDATA_FOLDER, "events.txt")

# ---------------- SETTINGS ----------------
settings = {
    "log_file": "",
    "success_sound_file": "",   # Empty = default
    "failure_sound_file": "",
    "show_notifications": False
}

success_trigger = "Successfully activated content on all platforms"
failure_triggers = [
    "LogValkyrieFortniteEditorLiveEdit: Verbose: FValkyrieFortniteEditorLiveEdit::ServerConnectionLost",
    "LogValkyrieRequestManagerEditor: Error:"
]

stop_thread = False
status_message = "Initializing..."
last_success_time = "Never"
last_failure_time = "Never"
icon = None  # Tray icon reference

# ---------------- RESOURCE PATH ----------------
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.dirname(os.path.abspath(__file__))  # <<<
    return os.path.join(base_path, relative_path)

ICON_PATH = resource_path(os.path.join("assets", "icon.ico"))

# ---------------- SETTINGS FUNCTIONS ----------------
def load_settings():
    global settings
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r') as f:
                settings.update(json.load(f))
        except Exception:
            print("⚠ Failed to load settings, using defaults.")

def save_settings():
    try:
        with open(SETTINGS_FILE, 'w') as f:
            json.dump(settings, f)
    except Exception as e:
        print(f"⚠ Failed to save settings: {e}")

# ---------------- EVENT LOGGING ----------------
def log_event(event_type, message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    entry = f"[{timestamp}] {event_type} - {message}\n"
    try:
        with open(EVENT_LOG_FILE, "a", encoding="utf-8") as f:
            f.write(entry)
    except Exception as e:
        print(f"⚠ Failed to write event log: {e}")

# ---------------- STARTUP TOGGLE ----------------
def get_startup_shortcut_path():
    startup_folder = os.path.join(
        os.getenv("APPDATA"),
        r"Microsoft\Windows\Start Menu\Programs\Startup"
    )
    shortcut_name = "UEFNPushNotifier.lnk"  # fixed name
    return os.path.join(startup_folder, shortcut_name)

def is_startup_enabled():
    return os.path.exists(get_startup_shortcut_path())

def toggle_startup(icon_obj, item):
    shortcut_path = get_startup_shortcut_path()
    if is_startup_enabled():
        try:
            os.remove(shortcut_path)
            print("Startup disabled.")
        except Exception as e:
            print(f"Failed to remove startup shortcut: {e}")
    else:
        create_startup_shortcut(shortcut_path)
        print("Startup enabled.")
    icon_obj.update_menu()

def create_startup_shortcut(shortcut_path):
    pythoncom.CoInitialize()
    shell_link = pythoncom.CoCreateInstance(
        shell.CLSID_ShellLink, None,
        pythoncom.CLSCTX_INPROC_SERVER, shell.IID_IShellLink
    )
    exe_path = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(sys.argv[0])
    
    # Add the --startup argument
    shell_link.SetPath(exe_path)
    shell_link.SetArguments("--startup")
    shell_link.SetWorkingDirectory(os.path.dirname(exe_path))

    persist_file = shell_link.QueryInterface(pythoncom.IID_IPersistFile)
    persist_file.Save(shortcut_path, 0)


# ---------------- LOG DETECTION ----------------
def find_log_file():
    base_path = os.path.expandvars(
        r"C:\Users\%USERNAME%\AppData\Local\UnrealEditorFortnite\Saved\Logs"
    )
    if not os.path.exists(base_path):
        return ""
    
    logs = glob.glob(os.path.join(base_path, "UnrealEditorFortnite*.log"))
    if not logs:
        return ""
    
    return max(logs, key=os.path.getmtime)

# ---------------- SOUND + NOTIFY ----------------
def play_sound(file_path):
    """
    Plays user-selected sound if available, 
    or the bundled default push_notif.wav,
    else falls back to SystemAsterisk.
    """
    # 1. User-selected file
    if file_path and file_path != "" and os.path.exists(file_path):
        winsound.PlaySound(file_path, winsound.SND_FILENAME | winsound.SND_ASYNC)
        return

    # 2. Bundled default sound
    fallback_path = resource_path(os.path.join("assets", "default_success.wav"))
    if os.path.exists(fallback_path):
        winsound.PlaySound(fallback_path, winsound.SND_FILENAME | winsound.SND_ASYNC)
        return
    else:
        notify("Didnt","Work")

    # 3. Windows fallback
    winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS | winsound.SND_ASYNC)



def notify(title, message):
    log_event("NOTIFICATION", title)
    try:
        toast = Notification(
            app_id="UEFN Push Notifier",  # Your app name here
            title=title,
            msg=message,
            icon=ICON_PATH  # Path to your .ico file
        )
        toast.set_audio(audio.Default, loop=False)
        toast.show()
    except Exception as e:
        print(f"Notification failed: {e}")
    except Exception:
        pass

# ---------------- RESET ----------------
def reset_settings(icon_obj, item):
    settings["success_sound_file"] = ""
    settings["failure_sound_file"] = ""
    settings["show_notifications"] = False
    
    save_settings()
    update_status("Settings reset.")
    icon_obj.update_menu()
    notify("UEFN Push Notifier", "Settings have been reset to default.")

# ---------------- STATUS ----------------
def update_status(msg=None):
    global status_message, icon
    if msg:
        status_message = msg
    if icon:
        icon.update_menu()

# ---------------- LOG MONITOR ----------------
def monitor_log():
    global stop_thread, status_message
    global last_success_time, last_failure_time
    last_inode = None
    f = None

    while not stop_thread:
        log_file = settings["log_file"]
        if not log_file or not os.path.exists(log_file):
            auto_log = find_log_file()
            if auto_log:
                settings["log_file"] = auto_log
                save_settings()
                update_status("Monitoring Log")
            else:
                update_status("Waiting for log...")
            time.sleep(2)
            continue

        try:
            current_inode = os.stat(log_file).st_ino
            if last_inode != current_inode:
                last_inode = current_inode
                if f:
                    f.close()
                f = open(log_file, 'r', encoding='utf-8', errors='ignore')
                f.seek(0, 2)
                update_status("Monitoring Log")
        except Exception as e:
            update_status(f"Error: {e}")
            time.sleep(2)
            continue

        line = f.readline()
        if not line:
            time.sleep(0.5)
            continue

        line_lower = line.lower()
        if success_trigger.lower() in line_lower:
            print("✅ Push complete detected.")
            last_success_time = datetime.now().strftime("%H:%M:%S")
            update_status("Monitoring: " + os.path.basename(log_file))
            log_event("SUCCESS", "Push complete detected")
            play_sound(settings.get("success_sound_file", ""))
            notify("✅ Push complete", "")
        elif any(trigger.lower() in line_lower for trigger in failure_triggers):
            print("❌ Push Error detected.")
            last_failure_time = datetime.now().strftime("%H:%M:%S")
            update_status("Error detected!")
            log_event("FAILURE", "Push Error")
            play_sound(settings.get("failure_sound_file", ""))
            notify("❌ Push error!", "")

    if f:
        f.close()

# ---------------- TRAY ICON ----------------
def on_exit(icon_obj, item):
    global stop_thread
    stop_thread = True
    icon_obj.stop()

def change_log_file(icon_obj, item):
    root = Tk()
    root.withdraw()
    logs_folder = os.path.expandvars(r"C:\Users\%USERNAME%\AppData\Local\UnrealEditorFortnite\Saved\Logs")
    file_path = select_file(
        title="Select Log File",
        filetypes=(("Log files", "*.log"), ("All files", "*.*")),
        initialdir=logs_folder
    )
    root.destroy()
    if file_path:
        settings["log_file"] = file_path
        save_settings()
        update_status("Monitoring: " + os.path.basename(file_path))

def change_success_sound(icon_obj, item):
    root = Tk()
    root.withdraw()
    file_path = select_file(
        title="Select Success Sound",
        filetypes=(("WAV files", "*.wav"),)
    )
    root.destroy()
    if file_path:
        settings["success_sound_file"] = file_path
        save_settings()
        play_sound(file_path)  # Test sound immediately
        icon_obj.update_menu()

def select_file(title="Select File", filetypes=(("All files", "*.*"),), initialdir=""):
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    options = {
        "title": title,
        "filetypes": filetypes
    }
    if initialdir:
        options["initialdir"] = initialdir

    file_path = filedialog.askopenfilename(**options)
    root.destroy()
    return file_path

def change_failure_sound(icon_obj, item):
    root = Tk()
    root.withdraw()
    file_path = select_file(
        title="Select Failure Sound",
        filetypes=(("WAV files", "*.wav"),)
    )
    root.destroy()
    if file_path:
        settings["failure_sound_file"] = file_path
        save_settings()
        play_sound(file_path)  # Test sound immediately
        icon_obj.update_menu()

def toggle_notifications(icon_obj, item):
    settings["show_notifications"] = not settings.get("show_notifications", False)
    if settings.get("show_notifications", True):
        play_sound(settings.get("success_sound_file", ""))
        notify("✅ Notifications Turned On", "This is what notifications look like.")
    save_settings()
    icon_obj.update_menu()

def open_event_log(icon_obj, item):
    log_path = EVENT_LOG_FILE
    if os.path.exists(log_path):
        subprocess.Popen(['notepad.exe', log_path])
    else:
        notify("Log file not created", "Nothing has happened yet, so the log file hasn't been created.")
        update_status("No event log found.")

def create_icon():
    status_label = lambda _: f"Status: {status_message}"
    last_notification_label = lambda _: f"Last Success: {last_success_time} / Failure: {last_failure_time}"

    success_sound_label = lambda _: (
        f"Success Sound: {os.path.basename(settings['success_sound_file']) if settings['success_sound_file'] else 'Default'}"
    )
    failure_sound_label = lambda _: (
        f"Failure Sound: {os.path.basename(settings['failure_sound_file']) if settings['failure_sound_file'] else 'Default'}"
    )
    log_label = lambda _: (
        f"Change Log File (Current: {os.path.basename(settings['log_file']) if settings['log_file'] else 'Auto'})"
    )
    startup_label = lambda _: f"Open on startup: {'✓' if is_startup_enabled() else '✗'}"
    notify_label = lambda _: f"Show notifications: {'✓' if settings.get('show_notifications', False) else '✗'}"

    settings_menu = pystray.Menu(
        item(log_label, change_log_file),
        item(success_sound_label, change_success_sound),
        item(failure_sound_label, change_failure_sound),
        item(notify_label, toggle_notifications),
        item(startup_label, toggle_startup),
        item("Reset Settings", reset_settings)
    )

    return pystray.Icon(
        "UEFN Push Notifier",
        Image.open(ICON_PATH),
        "UEFN Push Notifier",
        menu=pystray.Menu(
            item(status_label, None, enabled=False),
            item(last_notification_label, None, enabled=False),
            item("Settings", settings_menu),
            item("Open Event Log", open_event_log),
            item('Exit', on_exit)
        )
    )

# ---------------- MAIN ----------------
if __name__ == "__main__":
    launched_from_startup = "--startup" in sys.argv
    load_settings()

    if not launched_from_startup:
        notify("👋", "Program started and monitoring logs.")
    log_event("LAUNCHED", "UEFN Push Notifier Opened")

    # Auto-detect log if missing or invalid
    if not settings["log_file"] or not os.path.exists(settings["log_file"]):
        auto_log = find_log_file()
        if auto_log:
            settings["log_file"] = auto_log
            save_settings()
            status_message = "Monitoring: " + os.path.basename(auto_log)
        else:
            status_message = "Waiting for log..."

    # Start monitoring thread
    thread = threading.Thread(target=monitor_log, daemon=True)
    thread.start()

    # Must run tray icon in main thread
    icon = create_icon()
    icon.run()
