# Hidden Scheduled Task & Activity Logger

This script creates a hidden Windows scheduled task to run a specified executable and logs user activity via MongoDB and Cloudinary.

## Features
- Creates a hidden scheduled task that runs as **SYSTEM** with elevated privileges.
- Logs user **login/logout** events.
- Captures the **active window title** and takes **screenshots**.
- Uploads screenshots to **Cloudinary**.
- Logs events to **MongoDB**.

## Prerequisites
- **Windows OS**
- **Python 3.13** (or a supported version) for development
- **Internet connection** (for Cloudinary & MongoDB connectivity)

### Automatic Package Installation
The script includes an **auto-installer** that installs any missing dependencies at runtime. Required packages:

- `pyautogui`
- `pymongo`
- `psutil`
- `cloudinary`
- `pywin32`
- `pillow`
- `pyscreeze`

Alternatively, you can manually install them using:

```sh
py -m pip install -r requirements.txt
```

**`requirements.txt`**:
```
pyautogui
pymongo
psutil
cloudinary
pywin32
pillow
pyscreeze
```

## Usage
### Run Activity Logging Mode
To log user activity, simply run the script:

```sh
py python.py
```

This logs a user **login event** and captures activity every **10 seconds**.

### Run Scheduled Task Creation Mode
To create a scheduled task, pass the `task` argument followed by:

1. **Task Name**
2. **Full path to the executable**
3. _(Optional)_ Start time in **ISO format** (YYYY-MM-DDTHH:MM:SS)

#### Example:
```sh
py python.py task TestTask "C:\Path\to\myapp.exe" 2023-10-30T14:00:00
```

## Converting to .exe
To convert the script into a standalone executable, use **PyInstaller**:

```sh
py -m pip install pyinstaller
pyinstaller --onefile python.py
```

This creates a single executable inside the `dist` folder, including all necessary dependencies.

> **Note:** If the build process fails or encounters errors, try **disabling your antivirus** temporarily.
