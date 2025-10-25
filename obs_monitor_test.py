import time
from obsws_python import ReqClient
import json

# --- Configuration ---
# Please ensure these match your OBS WebSocket settings
HOST = "localhost"
PORT = 4455
PASSWORD = "Marana7ha" # Make sure this is correct
# ---

def run_monitor_test():
    """
    Connects to OBS, gets monitor geometry, and opens a test projector on each.
    """
    client = None
    try:
        print("Connecting to OBS WebSocket...")
        client = ReqClient(host=HOST, port=PORT, password=PASSWORD)
        print("Successfully connected to OBS.")
    except Exception as e:
        print(f"Error connecting to OBS: {e}")
        print("Please ensure OBS is running and the WebSocket server is enabled with the correct password.")
        return

    try:
        monitors = client.get_monitor_list().monitors
    except Exception as e:
        print(f"Error getting monitor list from OBS: {e}")
        client.disconnect()
        return

    print(f"\n--- Found {len(monitors)} Monitors ---")
    print("The script will now open a test projector on each monitor.")
    print("Use the coordinates (e.g., monitor_x, monitor_y) to update your config.json.\n")

    for i, monitor in enumerate(monitors):
        x = monitor.get('monitorPositionX')
        y = monitor.get('monitorPositionY')
        width = monitor.get('monitorWidth')
        height = monitor.get('monitorHeight')

        print(f"--> Testing Monitor Index: {i}")
        print(f"    Coordinates: (x: {x}, y: {y})")
        print(f"    Resolution: {width}x{height}")
        
        try:
            # Open a projector on this monitor
            client.send("OpenVideoMixProjector", {
                "videoMixType": "OBS_WEBSOCKET_VIDEO_MIX_TYPE_PROGRAM",
                "monitorIndex": i,
            })
            print(f"    A test projector should now be open on this monitor.")
            time.sleep(4) # Wait for user to see

        except Exception as e:
            print(f"    Error opening projector for index {i}: {e}")

    print("\n--- Test Complete ---")
    print("Please close all the opened 'Program (Projector)' windows manually.")

    if client:
        client.disconnect()

if __name__ == "__main__":
    run_monitor_test()
