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
    Connects to OBS and cycles through monitor indexes to identify them.
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

    # The projector type to open. "Program" is a good neutral choice.
    projector_type = "OBS_WEBSOCKET_VIDEO_MIX_TYPE_PROGRAM"

    # Number of monitors to test. We detected 5.
    num_monitors = 5

    print("\n--- Starting Monitor Index Test ---")
    print("A projector window will be opened for each monitor index.")
    print("Please note which monitor the projector appears on for each index.\n")

    for i in range(num_monitors):
        print(f"--> Testing Monitor Index: {i}")
        try:
            # Open the projector
            client.send(
                "OpenVideoMixProjector",
                {
                    "videoMixType": projector_type,
                    "monitorIndex": i,
                },
            )
            print(f"    Projector command sent for index {i}. Please observe your monitors.")

            # Wait for a few seconds for the next one
            time.sleep(3)

        except Exception as e:
            print(f"    Error opening projector for index {i}: {e}")
            print("    This could mean the index is invalid or OBS returned an error.")
            time.sleep(2)


    print("\n--- Test Complete ---")
    print("The script has finished cycling through all indexes.")
    print("IMPORTANT: Please close all the opened 'Program (Projector)' windows manually.")
    print("Once you have your list of which index corresponds to which monitor, you can update your config.json.")


    if client:
        client.disconnect()

if __name__ == "__main__":
    run_monitor_test()
