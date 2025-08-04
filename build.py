import os
import shutil
import subprocess
import re

# Path to your main Python file
MAIN_FILE = os.path.join("src", "uefn_push_notifier.py")

# Extract version from __version__ variable
with open(MAIN_FILE, "r", encoding="utf-8") as f:
    content = f.read()
match = re.search(r'__version__\s*=\s*["\'](.+?)["\']', content)
if not match:
    raise RuntimeError("Could not find __version__ in main script.")
VERSION = match.group(1)

# Create release folder
release_folder = os.path.join("releases", f"v{VERSION}")
os.makedirs(release_folder, exist_ok=True)

# Run PyInstaller
cmd = [
    "pyinstaller",
    "--onefile",
    "--windowed",
    f"--icon=assets/icon.ico",
    f"--add-data=assets;assets",
    "src/uefn_notifier.py"
]
subprocess.run(cmd, check=True)

# Move dist and build to release folder
for folder in ["dist", "build"]:
    if os.path.exists(folder):
        shutil.move(folder, os.path.join(release_folder, folder))

# Move the spec file
spec_file = "uefn_notifier.spec"
if os.path.exists(spec_file):
    shutil.move(spec_file, os.path.join(release_folder, spec_file))

print(f"✅ Build complete. Files moved to {release_folder}")
