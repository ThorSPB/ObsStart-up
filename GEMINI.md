# GEMINI.md

## Project Overview

This project contains a Python script (`obsStart.py`) designed to automate the startup and management of OBS (Open Broadcaster Software) projectors on Windows. It ensures that specific OBS projectors are always running, and it can also launch the OBSBOT Center application.

The script is highly configurable and can be set to run in two modes:
1.  **Single Run:** The script checks and opens the required projectors once and then exits.
2.  **Monitor Mode:** The script runs continuously, checking for missing projectors at a set interval and reopening them if they are closed.

### Key Technologies

*   **Python 3**
*   **OBS WebSocket:** The script communicates with OBS using the `obsws-python` library to control projectors.
*   **Windows API:** The script uses `pywin32` to interact with Windows for tasks like finding and focusing windows, and suppressing taskbar flashing.
*   **psutil:** Used to check if OBS and OBSBOT Center are running.

## Building and Running

### Prerequisites

*   Python 3
*   OBS Studio with the WebSocket server enabled.
*   The Python packages listed in `requirements.txt`.

### Installation

1.  **Clone the repository or download the files.**
2.  **Install the required Python packages:**
    ```bash
    pip install -r requirements.txt
    ```

### Configuration

The script uses a `config.json` file located in `%APPDATA%\ObsStartUp\`. If the file does not exist, the script will create a default one. You can edit this file to define the projectors you want to manage.

**Default `config.json`:**
```json
{
    "2": {"title": "Program (Projector)", "type": "program", "monitor": 3},
    "3": {"title": "Scene Projector (Proiector)", "type": "scene", "monitor": 1, "scene": "Proiector"},
    "4": {"title": "Scene Projector (TV Sala)", "type": "scene", "monitor": 4, "scene": "TV Sala"}
}
```

### Running the Script

To run the script, simply execute the `obsStart.py` file:

```bash
python obsStart.py
```

You can configure the script's behavior by changing the `MONITOR_MODE` variable at the top of the `obsStart.py` file.

*   `MONITOR_MODE = True`: Runs the script in continuous monitoring mode.
*   `MONITOR_MODE = False`: Runs the script as a single check.

## A Note on Monitor Indexing

A critical part of this script's functionality is opening projectors on specific monitors, which is defined by the `monitor` index in the `config.json` file.

It has been observed that OBS Studio does not always number monitors in a predictable way (e.g., simple left-to-right order). The index OBS uses may not match the order shown in Windows Display Settings or the physical arrangement of the monitors.

If you find that projectors are opening on the wrong screens, you must find the correct monitor index that OBS is using. This project now includes a diagnostic script to help with this.

**To find the correct monitor indexes:**
1.  Make sure OBS is running.
2.  Run the `obs_monitor_test.py` script:
    ```bash
    python obs_monitor_test.py
    ```
3.  The script will attempt to open a projector on each monitor index, starting from 0.
4.  Observe which monitor the projector appears on for each index and note the mapping.
5.  Update the `monitor` values in your `config.json` with the correct indexes you discovered.

## Development Conventions

*   The script is written in Python and follows standard Python conventions.
*   It makes heavy use of the `ctypes` and `win32` libraries for Windows-specific functionality.
*   Configuration is stored in a separate JSON file to keep it separate from the code.
*   The script includes detailed print statements to provide feedback on its progress and any errors that occur.
*   The PyInstaller spec file (`obsLauncher.spec`) is configured to use the `OBS_Studio_logo.ico` file for the final executable.
