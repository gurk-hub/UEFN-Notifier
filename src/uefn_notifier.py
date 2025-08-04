import os
import sys
import time
import winsound
import threading
import json
import glob
import subprocess
from datetime import datetime
from tkinter import Tk, filedialog, simpledialog, messagebox, ttk

import pystray
from pystray import MenuItem as item
from PIL import Image

import pythoncom
from win32com.shell import shell
from winotify import Notification, audio

# ---------------- PATHS ----------------
APPDATA_FOLDER = os.path.join(os.getenv("APPDATA"), "UEFNNotifier")
os.makedirs(APPDATA_FOLDER, exist_ok=True)

SETTINGS_FILE = os.path.join(APPDATA_FOLDER, "settings.json")
EVENT_LOG_FILE = os.path.join(APPDATA_FOLDER, "events.txt")

__version__ = "1.4.1"

DEFAULT_SETTINGS = {
    "log_file": "",
    "show_notifications": True,
    "triggers": [
        {
            "name": "✅ Session Connected",
            "keywords": ["EMemorySamplerState::Ready"],
            "sound_file": "default_success.wav",
            "notify": True
        },
        {
            "name": "❌ Push Failure",
            "keywords": [
                "LogValkyrieRequestManagerEditor: Error"
            ],
            "sound_file": "",
            "notify": True
        },
        {
            "name": "✅ HLOD Generated",
            "keywords": [
                "LogEditorBuildUtils: Build time"
            ],
            "sound_file": "default_success.wav",
            "notify": True
        },
        {
            "name": "❌ HLOD Failure",
            "keywords": [
                "LogWorldPartitionEditor: Error"
            ],
            "sound_file": "",
            "notify": True
        }
    ]
}

# Track all open windows
open_windows = []

stop_thread = False
status_message = "Initializing..."
last_trigger_time = "--:--:--"
icon = None  # Tray icon reference

# ---------------- RESOURCE PATH ----------------
def resource_path(relative_path: str) -> str:
    """Get absolute path to resource (works in PyInstaller onefile)."""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

ICON_PATH = resource_path(os.path.join("assets", "icon.ico"))

# ---------------- SETTINGS ----------------
def load_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r') as f:
                return json.load(f)
        except:
            print("⚠ Failed to load settings, using defaults.")
    return DEFAULT_SETTINGS.copy()

def save_settings():
    try:
        with open(SETTINGS_FILE, 'w') as f:
            json.dump(settings, f, indent=4)
    except Exception as e:
        print(f"⚠ Failed to save settings: {e}")

settings = load_settings()

# After loading settings
if "triggers" not in settings:
    settings["triggers"] = [
        {
            "name": "✅ Session Connected",
            "keywords": ["EMemorySamplerState::Ready"],
            "sound_file": "default_success.wav",
            "notify": True
        },
        {
            "name": "❌ Push Failure",
            "keywords": [
                "LogValkyrieRequestManagerEditor: Error"
            ],
            "sound_file": "",
            "notify": True
        },
        {
            "name": "✅ HLOD Generated",
            "keywords": [
                "LogEditorBuildUtils: Build time"
            ],
            "sound_file": "default_success.wav",
            "notify": True
        },
        {
            "name": "❌ HLOD Failure",
            "keywords": [
                "LogWorldPartitionEditor: Error"
            ],
            "sound_file": "",
            "notify": True
        }
    ]
    save_settings()

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
    return os.path.join(startup_folder, "UEFNNotifier.lnk")

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
    return max(logs, key=os.path.getmtime) if logs else ""

# ---------------- SOUND + NOTIFY ----------------
def play_sound(file_path):
    """Play a sound from either absolute path or assets folder."""
    if not file_path:
        winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS | winsound.SND_ASYNC)
        return

    # Absolute or asset path
    path = file_path if os.path.isabs(file_path) else resource_path(os.path.join("assets", file_path))
    if os.path.exists(path):
        winsound.PlaySound(path, winsound.SND_FILENAME | winsound.SND_ASYNC)
    else:
        winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS | winsound.SND_ASYNC)

def notify(title, message):
    log_event("NOTIFY", title)
    try:
        toast = Notification(
            app_id="UEFN Notifier",
            title=title,
            msg=message,
            icon=ICON_PATH
        )
        toast.set_audio(audio.Default, loop=False)
        toast.show()
    except:
        pass

# ---------------- RESET ----------------
def reset_settings(icon_obj=None, item=None):
    global settings
    settings = DEFAULT_SETTINGS.copy()
    save_settings()
    update_status("Settings reset.")
    if icon_obj: icon_obj.update_menu()
    notify("UEFN Notifier", "Settings have been reset to default.")

# ---------------- TRIGGER MANAGEMENT ----------------
def manage_triggers_gui():
    def refresh_tree():
        for row in tree.get_children():
            tree.delete(row)
        for i, trig in enumerate(settings.get("triggers", [])):
            keywords = ", ".join(trig.get("keywords", []))
            sound_file = trig.get("sound_file", "")
            sound_name = os.path.basename(sound_file) if sound_file else "Default"
            notify_status = "✓" if trig.get("notify", True) else "✗"

            tree.insert("", "end", iid=i, values=(trig.get("name", ""), keywords, sound_name, notify_status))

    def add_trigger():
        name = simpledialog.askstring("Add Trigger", "Enter trigger name:")
        if not name:
            return
        keywords = simpledialog.askstring("Add Trigger", "Enter keywords (comma-separated):")
        if not keywords:
            return
        sound_file = filedialog.askopenfilename(title="Select Sound File", filetypes=[("WAV files", "*.wav")])
        if not sound_file:
            return
        new_trigger = {
            "name": name,
            "keywords": [k.strip() for k in keywords.split(",")],
            "sound_file": sound_file,
            "notify": True
        }
        settings.setdefault("triggers", []).append(new_trigger)
        save_settings()
        refresh_tree()
        update_status(f"Trigger '{name}' added.")

    def edit_name_keywords():
        selected = tree.selection()
        if not selected:
            messagebox.showinfo("Edit Trigger", "Please select a trigger to edit.")
            return
        idx = tree.index(selected[0])
        trig = settings["triggers"][idx]

        new_name = simpledialog.askstring("Edit Name", "Enter new trigger name:", initialvalue=trig.get("name", ""))
        if not new_name:
            return

        new_keywords = simpledialog.askstring(
            "Edit Keywords", 
            "Enter keywords (comma-separated):", 
            initialvalue=", ".join(trig.get("keywords", []))
        )
        if not new_keywords:
            return

        trig["name"] = new_name
        trig["keywords"] = [k.strip() for k in new_keywords.split(",")]
        save_settings()
        refresh_tree()
        update_status(f"Trigger '{new_name}' updated.")

    def change_sound():
        selected = tree.selection()
        if not selected:
            messagebox.showinfo("Change Sound", "Please select a trigger to change sound.")
            return
        idx = tree.index(selected[0])
        trig = settings["triggers"][idx]

        sound_file = filedialog.askopenfilename(title="Select New Sound File", filetypes=[("WAV files", "*.wav")])
        if not sound_file:
            return

        trig["sound_file"] = sound_file
        save_settings()
        refresh_tree()
        update_status(f"Sound changed for trigger '{trig.get('name', '')}'.")

    def toggle_notify():
        selected = tree.selection()
        if not selected:
            messagebox.showwarning("No selection", "Please select a trigger to toggle notification.")
            return
        idx = int(selected[0])
        trig = settings["triggers"][idx]
        trig["notify"] = not trig.get("notify", True)
        save_settings()
        refresh_tree()
        update_status(f"Notification toggled for '{trig['name']}'")

    def delete_trigger():
        selected = tree.selection()
        if not selected:
            messagebox.showinfo("Delete Trigger", "Please select a trigger to delete.")
            return
        idx = tree.index(selected[0])
        trig = settings["triggers"][idx]

        confirm = messagebox.askyesno("Delete Trigger", f"Are you sure you want to delete trigger '{trig.get('name', '')}'?")
        if confirm:
            del settings["triggers"][idx]
            save_settings()
            refresh_tree()
            update_status(f"Trigger '{trig.get('name', '')}' deleted.")

    window = Tk()
    open_windows.append(window)
    window.title("Manage Triggers")
    window.geometry("600x400")
    #window.resizable(False, False)

    tree = ttk.Treeview(window, columns=("Name", "Keywords", "Sound", "Notify"), show="headings")
    tree.heading("Name", text="Name")
    tree.heading("Keywords", text="Keywords")
    tree.heading("Sound", text="Sound")
    tree.heading("Notify", text="Notify")
    tree.column("Name", width=150)
    tree.column("Keywords", width=260)
    tree.column("Sound", width=100)
    tree.column("Notify", width=50, anchor="center")
    tree.pack(fill="both", expand=True, padx=10, pady=10)

    btn_frame = ttk.Frame(window)
    btn_frame.pack(fill="x", padx=10, pady=(0,10))

    btn_add = ttk.Button(btn_frame, text="Add Trigger", command=add_trigger)
    btn_add.pack(side="left", padx=5)

    btn_edit = ttk.Button(btn_frame, text="Edit Name/Keywords", command=edit_name_keywords)
    btn_edit.pack(side="left", padx=5)

    btn_sound = ttk.Button(btn_frame, text="Change Sound", command=change_sound)
    btn_sound.pack(side="left", padx=5)

    btn_toggle_notify = ttk.Button(btn_frame, text="Toggle Notification", command=toggle_notify)
    btn_toggle_notify.pack(side="left", padx=5)

    btn_delete = ttk.Button(btn_frame, text="Delete Trigger", command=delete_trigger)
    btn_delete.pack(side="left", padx=5)

    btn_close = ttk.Button(btn_frame, text="Close", command=window.destroy)
    btn_close.pack(side="right", padx=5)

    refresh_tree()
    window.mainloop()

# ---------------- STATUS ----------------
def update_status(msg=None):
    global status_message, icon
    if msg: status_message = msg
    if icon: icon.update_menu()

# ---------------- LOG MONITOR ----------------
def monitor_log():
    """Continuously monitor the log file for trigger keywords."""
    global stop_thread, last_trigger_time

    last_inode = None
    f = None
    status_reset_timer = None

    def reset_status():
        """Reset the tray icon status after a trigger."""
        if not stop_thread:
            update_status("Monitoring Log")

    while not stop_thread:
        log_file = settings.get("log_file", "")
        
        # Auto-detect log file if missing
        if not log_file or not os.path.exists(log_file):
            auto_log = find_log_file()
            if auto_log:
                settings["log_file"] = auto_log
                save_settings()
                update_status("Monitoring log")
            else:
                update_status("Waiting for log...")
            time.sleep(2)
            continue

        try:
            # Detect if log file rotated / replaced
            current_inode = os.stat(log_file).st_ino
            if current_inode != last_inode:
                last_inode = current_inode
                if f:
                    f.close()
                f = open(log_file, 'r', encoding='utf-8', errors='ignore')
                f.seek(0, 2)  # Move to end
                update_status("Monitoring Log")
        except Exception as e:
            update_status(f"Error opening log: {e}")
            time.sleep(2)
            continue

        line = f.readline()
        if not line:
            time.sleep(0.5)
            continue

        # Check triggers
        line_lower = line.lower()
        triggered = False
        for trigger in settings.get("triggers", []):
            for keyword in trigger.get("keywords", []):
                if keyword.lower() in line_lower:
                    last_trigger_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    log_event(trigger["name"].upper(), f"Triggered by: {keyword}")

                    # Play sound
                    play_sound(trigger.get("sound_file", ""))

                    # Show notification if allowed
                    if settings.get("show_notifications", False) and trigger.get("notify", True):
                        notify(trigger["name"], f"{keyword}")

                    # Update tray status
                    update_status(f"Triggered: {trigger['name']}")
                    triggered = True

                    # Reset status after 5 seconds
                    if status_reset_timer:
                        status_reset_timer.cancel()
                    status_reset_timer = threading.Timer(5, reset_status)
                    status_reset_timer.start()
                    break  # Stop checking this line
            if triggered:
                break

    # Clean up on exit
    if f:
        f.close()

# ---------------- TRAY ICON ----------------
def on_exit(icon_obj, item):
    global stop_thread, thread
    stop_thread = True

    # Close all Tk windows to prevent hanging
    for w in open_windows:
        try:
            w.destroy()
        except:
            pass

    if thread.is_alive():
        thread.join(timeout=5)
    icon_obj.stop()

def select_file(title="Select File", filetypes=(("All files", "*.*"),), initialdir=""):
    root = Tk()
    open_windows.append(root)
    root.withdraw()
    file_path = filedialog.askopenfilename(title=title, filetypes=filetypes, initialdir=initialdir)
    root.destroy()
    return file_path

def toggle_notifications(icon_obj, item):
    settings["show_notifications"] = not settings.get("show_notifications", False)
    notify("✅Notifications Enabled", "This is what they look like")
    save_settings()
    icon_obj.update_menu()

def open_event_log(icon_obj, item):
    log_path = EVENT_LOG_FILE
    if os.path.exists(log_path):
        subprocess.Popen(['notepad.exe', log_path])
    else:
        notify("Log file not created", "Nothing has happened yet, so the log file hasn't been created.")
        update_status("No event log found.")

def open_settings_file(icon, item):
    if os.path.exists(SETTINGS_FILE):
        subprocess.Popen(['notepad.exe', SETTINGS_FILE])
    else:
        # Optionally notify or create a blank settings file first
        with open(SETTINGS_FILE, 'w') as f:
            f.write('{}')
        subprocess.Popen(['notepad.exe', SETTINGS_FILE])

def create_icon():
    status_label = lambda _: f"Status: {status_message}"
    last_label = lambda _: f"Last Trigger: {last_trigger_time}"
    startup_label = lambda _: f"Open On Startup: {'✓' if is_startup_enabled() else '✗'}"
    notify_label = lambda _: f"Show Notifications: {'✓' if settings.get('show_notifications', False) else '✗'}"

    settings_menu = pystray.Menu(
        item("Manage Triggers", lambda icon, item: manage_triggers_gui()),
        item(notify_label, toggle_notifications),
        item(startup_label, toggle_startup),
        item("Open Settings File", open_settings_file),
        item("Reset Settings", reset_settings)
    )

    return pystray.Icon(
        "UEFN Notifier",
        Image.open(ICON_PATH),
        "UEFN Notifier",
        menu=pystray.Menu(
            item(status_label, None, enabled=False),
            item(last_label, None, enabled=False),
            item("Settings", settings_menu),
            item("Open Event Log", open_event_log),
            item('Exit', on_exit)
        )
    )

# ---------------- MAIN ----------------
if __name__ == "__main__":
    launched_from_startup = "--startup" in sys.argv

    if not launched_from_startup:
        notify("👋", "Program started and monitoring logs.")
    log_event("LAUNCHED", "UEFN Notifier Opened")

    if not settings["log_file"] or not os.path.exists(settings["log_file"]):
        auto_log = find_log_file()
        if auto_log:
            settings["log_file"] = auto_log
            save_settings()
            status_message = "Monitoring: " + os.path.basename(auto_log)
        else:
            status_message = "Waiting for log..."

    thread = threading.Thread(target=monitor_log)
    thread.start()

    icon = create_icon()
    icon.run()
