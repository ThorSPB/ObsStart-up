#!/usr/bin/env python3
import time
import subprocess
import win32gui
import win32con
import win32api
import win32process
import psutil
import ctypes
from ctypes import wintypes
from obsws_python import ReqClient
import json
from monitor_utils import get_all_monitor_details

# Configuration - Verify these match your OBS setup
HOST = "localhost"
PORT = 4455
PASSWORD = "Marana7ha"
OBS_EXECUTABLE_PATH = r"C:\Program Files\obs-studio\bin\64bit\obs64.exe"  # Adjust path as needed
OBS_DIRECTORY = r"C:\Program Files\obs-studio\bin\64bit"  # OBS installation directory

# Monitoring settings
MONITOR_MODE = True  # Set to False for single run, True for continuous monitoring
CHECK_INTERVAL = 10  # Check every 10 seconds
STARTUP_DELAY = 20   # Wait 30 seconds after startup before first check

CONFIG = {}

# FlashWindowEx setup
FLASHW_STOP = 0
FLASHW_CAPTION = 1
FLASHW_TRAY = 2
FLASHW_ALL = 3

class FLASHWINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.UINT),
        ("hwnd",   wintypes.HWND),
        ("dwFlags", wintypes.DWORD),
        ("uCount", wintypes.UINT),
        ("dwTimeout", wintypes.DWORD),
    ]

def get_config_path():
    """Returns the path to the configuration file in AppData."""
    app_data = os.getenv('APPDATA')
    config_dir = os.path.join(app_data, "ObsStartUp")
    if not os.path.exists(config_dir):
        os.makedirs(config_dir)
    return os.path.join(config_dir, "config.json")

def load_config():
    """Loads the configuration from the JSON file, or creates it if it doesn't exist."""
    global CONFIG
    config_path = get_config_path()
    # NOTE: The configuration now uses monitor coordinates (e.g., 0, 1920) to identify
    # the target monitor. Use the obs_monitor_test.py script to find the correct coordinates.
    default_config = {
        "2": {"title": "Program (Projector)", "type": "program", "monitor_x": 0, "monitor_y": 0},
        "3": {"title": "Scene Projector (Proiector)", "type": "scene", "monitor_x": 1920, "monitor_y": 0, "scene": "Proiector"},
        "4": {"title": "Scene Projector (TV Sala)", "type": "scene", "monitor_x": -1920, "monitor_y": 0, "scene": "TV Sala"}
    }

    if os.path.exists(config_path):
        try:
            with open(config_path, 'r') as f:
                CONFIG = json.load(f)
            print(f"✅ Loaded configuration from {config_path}")
        except (json.JSONDecodeError, TypeError):
            print(f"⚠️ Invalid JSON in {config_path}. Using default config.")
            CONFIG = default_config
            with open(config_path, 'w') as f:
                json.dump(CONFIG, f, indent=4)
    else:
        print(f"📝 Configuration file not found. Creating default config at {config_path}")
        CONFIG = default_config
        with open(config_path, 'w') as f:
            json.dump(CONFIG, f, indent=4)

# Win32 constants for better window control
user32 = ctypes.windll.user32
ASFW_ANY = -1
VK_MENU = 0x12
KEYEVENTF_EXTENDEDKEY = 0x0001
KEYEVENTF_KEYUP = 0x0002


def get_monitors_sorted():
    """
    Gets all display monitors and returns them sorted by their horizontal position.
    """
    monitors = []
    MonitorEnumProc = ctypes.WINFUNCTYPE(
        ctypes.c_int,
        ctypes.c_ulong,
        ctypes.c_ulong,
        ctypes.POINTER(wintypes.RECT),
        ctypes.c_double
    )

    class MONITORINFOEXW(ctypes.Structure):
        _fields_ = [
            ("cbSize", wintypes.DWORD),
            ("rcMonitor", wintypes.RECT),
            ("rcWork", wintypes.RECT),
            ("dwFlags", wintypes.DWORD),
            ("szDevice", wintypes.WCHAR * 32)
        ]

    def enum_proc(hMonitor, hdcMonitor, lprcMonitor, dwData):
        info = MONITORINFOEXW()
        info.cbSize = ctypes.sizeof(MONITORINFOEXW)
        if ctypes.windll.user32.GetMonitorInfoW(hMonitor, ctypes.byref(info)):
            monitors.append({
                'hMonitor': hMonitor,
                'rcMonitor': info.rcMonitor,
                'szDevice': info.szDevice
            })
        return 1

    # Enumerate all display monitors
    ctypes.windll.user32.EnumDisplayMonitors(0, 0, MonitorEnumProc(enum_proc), 0)

    # Sort monitors by their left coordinate
    return sorted(monitors, key=lambda m: m['rcMonitor'].left)


def get_monitor_index_from_coords(x, y, client):
    """
    Finds the OBS monitor index that corresponds to a given (x, y) coordinate.
    """
    try:
        # Get the list of monitors from OBS
        monitors = client.get_monitor_list().monitors

        for i, monitor in enumerate(monitors):
            if monitor.get('monitorPositionX') == x and monitor.get('monitorPositionY') == y:
                print(f"  ✅ Found monitor index {i} for coordinates ({x}, {y})")
                return i
        
        print(f"  ⚠️ No monitor found in OBS for coordinates ({x}, {y}). Falling back to primary.")
        return 0 # Default to primary if not found
        
    except Exception as e:
        print(f"  ❌ Error getting monitor index from OBS: {e}")
        print("  Falling back to primary monitor (index 0).")
        return 0



def get_primary_monitor_rect():
    """
    Finds the rectangle of the primary display monitor.
    """
    MonitorEnumProc = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_ulong, ctypes.c_ulong, ctypes.POINTER(wintypes.RECT), ctypes.c_double)

    class MONITORINFOEXW(ctypes.Structure):
        _fields_ = [
            ("cbSize", wintypes.DWORD),
            ("rcMonitor", wintypes.RECT),
            ("rcWork", wintypes.RECT),
            ("dwFlags", wintypes.DWORD),
            ("szDevice", wintypes.WCHAR * 32)
        ]

    primary_rect = None
    def enum_proc(hMonitor, hdcMonitor, lprcMonitor, dwData):
        nonlocal primary_rect
        info = MONITORINFOEXW()
        info.cbSize = ctypes.sizeof(MONITORINFOEXW)
        if ctypes.windll.user32.GetMonitorInfoW(hMonitor, ctypes.byref(info)):
            if info.dwFlags & 1: # MONITORINFOF_PRIMARY
                primary_rect = info.rcMonitor
                return 0 # Stop enumeration
        return 1

    ctypes.windll.user32.EnumDisplayMonitors(0, 0, MonitorEnumProc(enum_proc), 0)
    return primary_rect

def check_and_correct_projector_positions(client):
    """
    Verifies that projectors are on the correct monitor and closes them if they are misplaced.
    """
    print("\n\U0001f50d Verifying projector positions...")
    
    try:
        obs_monitors = client.get_monitor_list().monitors
    except Exception as e:
        print(f"  \u26a0\ufe0f Could not get monitor list from OBS: {e}. Skipping position check.")
        return

    open_projectors = get_obs_projector_windows()
    if not open_projectors:
        return # Nothing to check

    for config_key, config in CONFIG.items():
        target_x = config.get('monitor_x')
        target_y = config.get('monitor_y')

        # Find the OBS monitor that matches the configured coordinates
        target_monitor_geom = None
        for monitor in obs_monitors:
            if monitor.get('monitorPositionX') == target_x and monitor.get('monitorPositionY') == target_y:
                target_monitor_geom = monitor
                break
        
        if not target_monitor_geom:
            print(f"  \u26a0\ufe0f No monitor found in OBS for coordinates ({target_x}, {target_y}) for '{config['title']}'.")
            continue

        # Find the corresponding window for this config entry
        found_hwnd = None
        for proj_window in open_projectors:
            title_lower = proj_window['title'].lower()
            is_match = False
            if config["type"] == "program" and "program" in title_lower:
                is_match = True
            elif config["type"] == "scene":
                scene_name = config["scene"].lower()
                if scene_name in title_lower or scene_name.replace(" ", "") in title_lower.replace(" ", ""):
                    is_match = True
            
            if is_match:
                found_hwnd = proj_window['hwnd']
                break
        
        if found_hwnd:
            # We found the window, now check its position.
            try:
                window_rect = win32gui.GetWindowRect(found_hwnd)
                window_center_x = (window_rect[0] + window_rect[2]) / 2
                window_center_y = (window_rect[1] + window_rect[3]) / 2

                # Check if the window's center is inside the target monitor's geometry
                mon_x = target_monitor_geom['monitorPositionX']
                mon_y = target_monitor_geom['monitorPositionY']
                mon_width = target_monitor_geom['monitorWidth']
                mon_height = target_monitor_geom['monitorHeight']

                if not (mon_x <= window_center_x < mon_x + mon_width and \
                        mon_y <= window_center_y < mon_y + mon_height):
                    
                    print(f"  \u26a0\ufe0f Misplaced projector detected: '{config['title']}' is not on the correct monitor.")
                    print(f"  Closing '{config['title']}' so it can be reopened correctly.")
                    win32gui.PostMessage(found_hwnd, win32con.WM_CLOSE, 0, 0)

            except Exception as e:
                print(f"  \u26a0\ufe0f Could not verify position for '{config['title']}': {e}")


def is_obs_running():
    """Check if OBS is already running and responsive"""
    try:
        for proc in psutil.process_iter(['name', 'status']):
            proc_name = proc.info['name'].lower()
            if 'obs64.exe' in proc_name or 'obs.exe' in proc_name:
                try:
                    if proc.status() in [psutil.STATUS_RUNNING, psutil.STATUS_SLEEPING]:
                        return True
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
        return False
    except Exception:
        return False


import os

def remove_obs_safe_mode_flag():
    """Removes the OBS 'safe_mode' file to suppress safe mode prompt."""
    safe_mode_file = os.path.expanduser(r"~\AppData\Roaming\obs-studio\safe_mode")
    if os.path.exists(safe_mode_file):
        try:
            os.remove(safe_mode_file)
            print("🧹 Removed 'safe_mode' file")
        except Exception as e:
            print(f"⚠️ Couldn't remove 'safe_mode': {e}")
    else:
        print("✅ No 'safe_mode' file present")

def focus_window(hwnd):
    """Restore and focus hwnd reliably, even across processes."""
    try:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        user32.AllowSetForegroundWindow(ASFW_ANY)
        win32api.keybd_event(VK_MENU, 0, KEYEVENTF_EXTENDEDKEY, 0)
        win32api.keybd_event(VK_MENU, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0)
        win32gui.SetForegroundWindow(hwnd)
        win32gui.SetActiveWindow(hwnd)
        win32gui.BringWindowToTop(hwnd)
        user32.SwitchToThisWindow(hwnd, True)
        print("✅ Focused window:", hwnd)
        return True
    except Exception as e:
        print(f"⚠️ Could not focus window {hwnd}: {e}")
        return False

def suppress_taskbar_flash_aggressive(hwnd, max_attempts=3):
    """Aggressively suppress taskbar flash with multiple strategies"""
    try:
        for attempt in range(max_attempts):
            try:
                # Stop any current flashing immediately - multiple methods
                fwi = FLASHWINFO(ctypes.sizeof(FLASHWINFO), hwnd, FLASHW_STOP, 0, 0)
                ctypes.windll.user32.FlashWindowEx(ctypes.byref(fwi))
                
                # Alternative flash stop method
                ctypes.windll.user32.FlashWindow(hwnd, False)
                
                # Change window extended styles to hide from taskbar
                try:
                    ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
                    new_ex = (ex_style & ~win32con.WS_EX_APPWINDOW) | win32con.WS_EX_TOOLWINDOW
                    win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, new_ex)
                except:
                    pass  # Sometimes this fails, continue anyway
                
                # Force window position without activation to prevent flash
                win32gui.SetWindowPos(
                    hwnd, win32con.HWND_BOTTOM,
                    0, 0, 0, 0,
                    win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE
                )
                
                # Small delay
                time.sleep(0.05)
                
                # Bring back to top without activation
                win32gui.SetWindowPos(
                    hwnd, win32con.HWND_TOP,
                    0, 0, 0, 0,
                    win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE | win32con.SWP_SHOWWINDOW
                )
                
                # Final flash suppression
                fwi = FLASHWINFO(ctypes.sizeof(FLASHWINFO), hwnd, FLASHW_STOP, 0, 0)
                ctypes.windll.user32.FlashWindowEx(ctypes.byref(fwi))
                ctypes.windll.user32.FlashWindow(hwnd, False)
                
                if attempt == 0:
                    print(f"  🔇 Flash suppression applied to window {hwnd}")
                break
                
            except Exception as e:
                if attempt < max_attempts - 1:
                    time.sleep(0.1)
                    continue
                else:
                    print(f"  ⚠️ Flash suppression partially failed for {hwnd}: {e}")
                    
    except Exception as e:
        print(f"⚠️ Could not suppress flash for hwnd={hwnd}: {e}")

def start_obs():
    """Start OBS if it's not already running"""
    if is_obs_running():
        print("✅ OBS is already running")
        return True

    remove_obs_safe_mode_flag()
    
    print("🚀 Starting OBS...")
    try:
        subprocess.Popen([OBS_EXECUTABLE_PATH, "--disable-safe-mode"], cwd=OBS_DIRECTORY, shell=False)
        print("✅ OBS started successfully")
        
        print("⏳ Waiting for OBS to initialize...")
        time.sleep(4)

        # Find and focus OBS main window
        def find_obs_main_window():
            def callback(hwnd, extra):
                if win32gui.IsWindowVisible(hwnd):
                    title = win32gui.GetWindowText(hwnd)
                    class_name = win32gui.GetClassName(hwnd)
                    if ("OBS" in title or "obs64" in title.lower()) and "Qt" in class_name:
                        extra.append(hwnd)
                return True
            hwnds = []
            win32gui.EnumWindows(callback, hwnds)
            return hwnds[0] if hwnds else None

        hwnd = find_obs_main_window()
        if hwnd:
            try:
                if not focus_window(hwnd):
                    # Fallback focusing method
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    foreground_hwnd = win32gui.GetForegroundWindow()
                    current_thread_id = win32api.GetCurrentThreadId()
                    foreground_thread_id = win32process.GetWindowThreadProcessId(foreground_hwnd)[0]
                    win32api.AttachThreadInput(current_thread_id, foreground_thread_id, True)
                    win32gui.SetForegroundWindow(hwnd)
                    win32gui.BringWindowToTop(hwnd)
                    win32api.AttachThreadInput(current_thread_id, foreground_thread_id, False)
                print("✅ OBS window focused")
            except Exception as e:
                print(f"⚠️ Could not focus OBS window: {e}")
        else:
            print("⚠️ OBS window not found to focus")

        if is_obs_running():
            print("✅ OBS is now running")
            return True
        else:
            print("❌ OBS failed to start properly")
            return False
            
    except FileNotFoundError:
        print(f"❌ OBS executable not found at: {OBS_EXECUTABLE_PATH}")
        print("💡 Please update OBS_EXECUTABLE_PATH in the script")
        return False
    except Exception as e:
        print(f"❌ Failed to start OBS: {e}")
        return False

def connect_to_obs_websocket(max_retries=5):
    """Connect to OBS WebSocket with retries"""
    for attempt in range(max_retries):
        try:
            client = ReqClient(host=HOST, port=PORT, password=PASSWORD)
            print("✅ Connected to OBS WebSocket")
            return client
        except Exception as e:
            print(f"⏳ WebSocket connection attempt {attempt + 1}/{max_retries} failed: {e}")
            if attempt < max_retries - 1:
                time.sleep(3)
            else:
                print("❌ Failed to connect to OBS WebSocket after all retries")
                print("💡 Make sure OBS WebSocket server is enabled in OBS settings")
                return None

def get_obs_projector_windows():
    """Get all OBS projector windows with their handles"""
    projectors = []
    
    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            class_name = win32gui.GetClassName(hwnd)
            
            if ("Projector" in title and 
                ("OBS" in title or "obs64" in class_name.lower() or "Qt" in class_name)):
                projectors.append({
                    'hwnd': hwnd,
                    'title': title,
                    'class': class_name
                })
        return True
    
    win32gui.EnumWindows(callback, None)
    return projectors

def wait_for_projector_window(config, timeout=8):
    """Wait for a specific projector window to appear and return its handle"""
    start_time = time.time()
    
    while time.time() - start_time < timeout:
        projectors = get_obs_projector_windows()
        
        for proj in projectors:
            title_lower = proj['title'].lower()
            
            if config["type"] == "program" and "program" in title_lower:
                return proj["hwnd"]
            elif config["type"] == "scene":
                scene_name = config["scene"].lower()
                if scene_name in title_lower or scene_name.replace(" ", "") in title_lower.replace(" ", ""):
                    return proj["hwnd"]
        
        time.sleep(0.2)
    
    return None

def open_projector_with_flash_suppression(client, config, monitor_details):
    """
    Open a single projector and immediately suppress its taskbar flash.
    Returns:
        - True: If the projector was opened successfully.
        - False: If there was an error during the process.
        - None: If the projector was skipped because the monitor is off.
    """
    try:
        target_x = config.get('monitor_x', 0)
        target_y = config.get('monitor_y', 0)

        # Check if the target monitor is active before trying to open it.
        target_monitor = next((m for m in monitor_details if m['rect'].left == target_x and m['rect'].top == target_y), None)
        
        # If monitor is found and not active, skip it.
        if target_monitor and not target_monitor['is_active']:
            print(f"  💤 Skipping '{config['title']}' because monitor at ({target_x}, {target_y}) is off or in power-save mode.")
            return None  # Special return value for "skipped"

        monitor_index = get_monitor_index_from_coords(target_x, target_y, client)

        # Open the projector
        if config["type"] == "program":
            client.send("OpenVideoMixProjector", {
                "videoMixType": "OBS_WEBSOCKET_VIDEO_MIX_TYPE_PROGRAM",
                "monitorIndex": monitor_index
            })
            print(f"  📺 Opening Program projector on monitor {monitor_index}")
            
        elif config["type"] == "scene":
            client.send("OpenSourceProjector", {
                "sourceName": config["scene"],
                "monitorIndex": monitor_index
            })
            print(f"  📺 Opening {config['scene']} projector on monitor {monitor_index}")
        
        hwnd = wait_for_projector_window(config, timeout=6)
        
        if hwnd:
            suppress_taskbar_flash_aggressive(hwnd, max_attempts=1)
            return True
        else:
            print(f"  ⚠️ Could not find window handle for {config['title']}")
            return False
            
    except Exception as e:
        print(f"  ❌ Failed to open {config['title']}: {e}")
        return False

def check_missing_projectors():
    """Check which projectors are missing and which exist"""
    existing = get_obs_projector_windows()
    missing = []
    found = []
    
    for monitor, config in CONFIG.items():
        projector_found = False
        for proj in existing:
            title_lower = proj['title'].lower()
            
            if config["type"] == "program" and "program" in title_lower:
                projector_found = True
                found.append(monitor)
                break
            elif config["type"] == "scene":
                scene_name = config["scene"].lower()
                if scene_name in title_lower or scene_name.replace(" ", "") in title_lower.replace(" ", ""):
                    projector_found = True
                    found.append(monitor)
                    break
        
        if not projector_found:
            missing.append(monitor)
    
    return missing, found

def open_missing_projectors_enhanced(client):
    """Enhanced version with better flash suppression and monitor status check"""
    print("💻 Checking monitor power states...")
    try:
        monitor_details = get_all_monitor_details()
        active_monitors_count = sum(1 for m in monitor_details if m['is_active'])
        print(f"  → Found {len(monitor_details)} total monitors, {active_monitors_count} are active.")
    except Exception as e:
        print(f"  ⚠️ Could not check monitor power states: {e}")
        print("     (Is the 'wmi' package installed? Falling back to basic check.)")
        monitor_details = [] # Fallback to empty list

    missing, found = check_missing_projectors()
    
    if not missing:
        print("✅ All required projectors are already running!")
        return True # Indicates no action was needed, but it's not a failure.
    
    print(f"📋 Found projectors: {found}")
    print(f"🔍 Missing projectors: {missing}")
    print("🚀 Opening missing projectors with flash suppression:")
    
    any_opened = False
    
    for monitor_id in missing:
        config = CONFIG[monitor_id]
        
        # In case WMI failed, create a dummy entry that will always be 'active'
        if not monitor_details:
            monitor_details = [{'rect': type('obj', (object,), {'left': config.get('monitor_x', 0), 'top': config.get('monitor_y', 0)})(), 'is_active': True}]

        result = open_projector_with_flash_suppression(client, config, monitor_details)
        if result is True:
            any_opened = True
            time.sleep(0.2) # Stagger opening projectors
    
    return any_opened

def verify_projectors_exist():
    """Check if projectors are actually running"""
    print("\n🔍 Verifying projectors:")
    
    projectors = get_obs_projector_windows()
    found_configs = []
    
    for proj in projectors:
        title = proj['title']
        print(f"  → Found window: {title}")
        
        title_lower = title.lower()
        if "program" in title_lower:
            found_configs.append(2)
        elif "proiector" in title_lower:
            found_configs.append(3)
        elif "tv sala" in title_lower:
            found_configs.append(4)
        
        for monitor, config in CONFIG.items():
            if config["type"] == "scene" and config["scene"].lower() in title_lower:
                if monitor not in found_configs:
                    found_configs.append(monitor)
    
    found_configs = list(set(found_configs))
    success = len(found_configs) >= len(CONFIG) - 1
    print(f"  → Expected: {len(CONFIG)}, Found matching: {len(found_configs)} {found_configs}")
    
    return success, projectors

def monitor_projectors_continuously():
    """Continuously monitor and maintain projectors"""
    print(f"🛡️ Starting continuous monitoring mode (checking every {CHECK_INTERVAL} seconds)")
    print("💡 This will run in the background and auto-recover any closed projectors")
    print("💡 If you close OBS, the script will automatically stop monitoring")
    print("💡 Press Ctrl+C to stop monitoring manually\n")
    
    if STARTUP_DELAY > 0:
        print(f"⏳ Startup delay: waiting {STARTUP_DELAY} seconds before first check...")
        time.sleep(STARTUP_DELAY)
    
    client = None
    check_count = 1
    
    try:
        while True:
            print(f"\n🔍 Monitor Check #{check_count} - {time.strftime('%H:%M:%S')}")
            
            if not is_obs_running():
                print("🛑 OBS has been closed - stopping monitoring gracefully")
                print("💡 To restart OBS and monitoring, run the script again")
                break
            else:
                print("🚀 OBS still running")

            if client is None:
                client = connect_to_obs_websocket(max_retries=2)
                if not client:
                    print("❌ WebSocket connection failed, will retry next cycle")
                    time.sleep(CHECK_INTERVAL)
                    check_count += 1
                    continue
            
            try:
                # Get current monitor and projector status at the start of each check
                monitor_details = get_all_monitor_details()
                missing, found = check_missing_projectors()
                
                if not missing:
                    print("✅ All projectors running correctly")
                else:
                    print(f"⚠️ Missing projectors detected: {missing}")
                    print(f"📋 Currently running: {found}")
                    
                    if not is_obs_running():
                        print("🛑 OBS was closed during check - stopping monitoring gracefully")
                        break
                    
                    for monitor_id in missing:
                        config = CONFIG[monitor_id]
                        # This function now uses the monitor details to decide whether to act
                        result = open_projector_with_flash_suppression(client, config, monitor_details)
                        
                        if result is True:
                            print(f"  ✅ Recovered {config['title']}")
                            time.sleep(0.5)
                        elif result is False:
                            # A hard error occurred, likely a connection issue.
                            print("  ❌ An error occurred. Will try to reconnect on the next cycle.")
                            client = None # Force re-connection
                            break # End this check cycle and start a new one
                
                # If the client connection was dropped, skip the position check for this cycle
                if client is None:
                    check_count += 1
                    time.sleep(CHECK_INTERVAL)
                    continue

                # After attempting to open missing projectors, verify all positions.
                time.sleep(1) # Give windows a moment to appear and settle.
                check_and_correct_projector_positions(client)

            except Exception as e:
                print(f"❌ Error during projector check: {e}")
                client = None # Force re-connection on next loop
            
            check_count += 1
            time.sleep(CHECK_INTERVAL)
            
    except KeyboardInterrupt:
        print("\n🛑 Monitoring stopped by user")
    except Exception as e:
        print(f"\n💥 Monitor crashed: {e}")
    finally:
        if client:
            try:
                client.disconnect()
            except:
                pass
        print("🔚 Monitoring ended")

def is_obsbot_running():
    """Check if OBSBOT Center is already running"""
    for proc in psutil.process_iter(['name']):
        try:
            if proc.info['name'] and 'obsbot' in proc.info['name'].lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return False

def run_single_check():
    """Run a single check and exit"""
    print("🎬 OBS Projector Auto-Manager")
    print("=" * 50)
    
    if not start_obs():
        print("\n💥 FAILURE: Could not start OBS")
        input("Press Enter to exit...")
        return
    
    client = connect_to_obs_websocket()
    if not client:
        print("\n💥 FAILURE: Could not connect to OBS")
        input("Press Enter to exit...")
        return
    
    try:
        print("\n🔍 Checking existing projectors...")
        if open_missing_projectors_enhanced(client):
            time.sleep(2)
            success, projectors = verify_projectors_exist()

            # Verify that projectors are on the correct monitors.
            time.sleep(1)
            check_and_correct_projector_positions(client)

            if success:
                print("\n\U0001f308 SUCCESS: All required projectors are now running!")
            else:
                print(f"\n\u26a0\ufe0f CHECK NEEDED: {len(projectors)} projectors currently running")
                print("\U0001f4a1 Some projectors might not have opened properly")
        else:
            print("\n💥 FAILURE: Could not open missing projectors")
            
    except Exception as e:
        print(f"\n💥 UNEXPECTED ERROR: {e}")
        
    finally:
        try:
            client.disconnect()
        except:
            pass

    # Launch OBSBOT Center after projectors are opened if not already running
    obsbot_shortcut = r"C:\Users\Public\Desktop\OBSBOT Center.lnk"
    if not is_obsbot_running():
        try:
            os.startfile(obsbot_shortcut)
            print("🚀 Launched OBSBOT Center")
        except Exception as e:
            print(f"❌ Failed to launch OBSBOT Center: {e}")
    else:
        print("\nℹ️ OBSBOT Center is already running, not launching another instance")

    print("\n✅ Script completed!\n")

def main():
    """Main function - chooses between single run or continuous monitoring"""
    load_config()
    if MONITOR_MODE:
        run_single_check()
        monitor_projectors_continuously()
    else:
        run_single_check()

if __name__ == "__main__":
    main()