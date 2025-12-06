import ctypes
from ctypes import wintypes
import wmi

# --- Ctypes-based PnPDeviceID retrieval ---

# Constants
CCHDEVICENAME = 32
CCHDEVICESTRING = 128
EDD_GET_DEVICE_INTERFACE_NAME = 0x00000001

# Structures
class RECT(wintypes.RECT):
    def __repr__(self):
        return f"RECT(left={self.left}, top={self.top}, right={self.right}, bottom={self.bottom})"

class MONITORINFOEXW(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("rcMonitor", RECT),
        ("rcWork", RECT),
        ("dwFlags", wintypes.DWORD),
        ("szDevice", wintypes.WCHAR * CCHDEVICENAME),
    ]
    def __init__(self, *args, **kw):
        super().__init__(*args, **kw)
        self.cbSize = ctypes.sizeof(self.__class__)

class DISPLAY_DEVICEW(ctypes.Structure):
    _fields_ = [
        ("cb", wintypes.DWORD),
        ("DeviceName", wintypes.WCHAR * CCHDEVICENAME),
        ("DeviceString", wintypes.WCHAR * CCHDEVICESTRING),
        ("StateFlags", wintypes.DWORD),
        ("DeviceID", wintypes.WCHAR * CCHDEVICESTRING),
        ("DeviceKey", wintypes.WCHAR * CCHDEVICESTRING),
    ]
    def __init__(self, *args, **kw):
        super().__init__(*args, **kw)
        self.cb = ctypes.sizeof(self.__class__)

# Function prototypes
user32 = ctypes.WinDLL("user32", use_last_error=True)
MONITORENUMPROC = ctypes.WINFUNCTYPE(
    wintypes.BOOL, wintypes.HMONITOR, wintypes.HDC, ctypes.POINTER(RECT), wintypes.LPARAM
)
user32.EnumDisplayMonitors.argtypes = [wintypes.HDC, ctypes.POINTER(RECT), MONITORENUMPROC, wintypes.LPARAM]
user32.EnumDisplayMonitors.restype = wintypes.BOOL
user32.GetMonitorInfoW.argtypes = [wintypes.HMONITOR, ctypes.POINTER(MONITORINFOEXW)]
user32.GetMonitorInfoW.restype = wintypes.BOOL
user32.EnumDisplayDevicesW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, ctypes.POINTER(DISPLAY_DEVICEW), wintypes.DWORD]
user32.EnumDisplayDevicesW.restype = wintypes.BOOL

def _get_pnp_id(hmonitor):
    """Internal function to retrieve PnPDeviceID for a given HMONITOR."""
    try:
        monitor_info = MONITORINFOEXW()
        if not user32.GetMonitorInfoW(hmonitor, ctypes.byref(monitor_info)):
            return None
        
        monitor_device_name = monitor_info.szDevice
        
        iDevNum = 0
        display_device = DISPLAY_DEVICEW()
        while user32.EnumDisplayDevicesW(None, iDevNum, ctypes.byref(display_device), 0):
            if display_device.DeviceName == monitor_device_name:
                jDevNum = 0
                monitor_display_device = DISPLAY_DEVICEW()
                while user32.EnumDisplayDevicesW(display_device.DeviceName, jDevNum, ctypes.byref(monitor_display_device), EDD_GET_DEVICE_INTERFACE_NAME):
                    if monitor_display_device.DeviceID:
                        return monitor_display_device.DeviceID
                    jDevNum += 1
            iDevNum += 1
    except Exception:
        return None
    return None

# --- WMI and main logic ---

def get_all_monitor_details():
    """
    Retrieves a detailed list of all monitors, including their coordinates, PNP ID, and power state. 
    
    Returns:
        A list of dictionaries, where each dict represents a monitor.
        Returns an empty list if there's an error.
        Example:
        [
            {
                'hMonitor': 12345,
                'rect': <RECT object>,
                'pnp_id': 'DISPLAY\\...', 
                'is_active': True
            },
            ...
        ]
    """
    monitors_details = []
    monitor_handles = []

    # 1. Enumerate monitors to get handles and rects
    def callback(hMonitor, hdcMonitor, lprcMonitor, dwData):
        monitor_handles.append(hMonitor)
        return True
    
    if not user32.EnumDisplayMonitors(None, None, MONITORENUMPROC(callback), 0):
        return []

    # 2. Get WMI monitor statuses
    wmi_statuses = {}
    try:
        w = wmi.WMI(namespace="root\cimv2")
        for monitor in w.Win32_DesktopMonitor():
            # Availability=3 means "Running/Full Power"
            wmi_statuses[monitor.PNPDeviceID] = (monitor.Availability == 3)
    except Exception:
        # WMI might fail, so we proceed without statuses
        pass

    # 3. Combine all information
    for hmon in monitor_handles:
        info = MONITORINFOEXW()
        if user32.GetMonitorInfoW(hmon, ctypes.byref(info)):
            pnp_id = _get_pnp_id(hmon)
            is_active = wmi_statuses.get(pnp_id, True) # Default to True if WMI fails or monitor not found

            monitors_details.append({
                'hMonitor': hmon,
                'rect': info.rcMonitor,
                'pnp_id': pnp_id,
                'is_active': is_active
            })
            
    return monitors_details

if __name__ == '__main__':
    # Example usage:
    details = get_all_monitor_details()
    if not details:
        print("Could not retrieve monitor details.")
    else:
        print(f"Found {len(details)} monitors.")
        for i, monitor in enumerate(details):
            print(f"--- Monitor {i+1} ---")
            print(f"  Handle: {monitor['hMonitor']}")
            print(f"  Coordinates: {monitor['rect']}")
            print(f"  PnP ID: {monitor['pnp_id']}")
            print(f"  Is Active: {monitor['is_active']}")
