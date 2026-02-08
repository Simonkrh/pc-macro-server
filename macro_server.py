from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import os
import subprocess
import pyautogui
import comtypes
from pycaw.pycaw import (
    AudioUtilities,
    ISimpleAudioVolume,
    IAudioEndpointVolume,
    IAudioMeterInformation,
)
from ctypes import cast, POINTER, c_float
from comtypes import CLSCTX_ALL
import ctypes
import base64
from PIL import Image
import win32gui
import win32ui
import io
import json
from threading import Lock
import warnings
from pycaw.constants import DEVICE_STATE, EDataFlow
from pycaw.pycaw import AudioUtilities
import keyboard
import re
import time
import shlex
import traceback

app = Flask(__name__)
CORS(app)

COMMAND_PATTERN = re.compile(
    r"<(enter|wait:\d+)>"
)  # Commands for type_text macro endpoint

# Macros
MACROS_FILE = "macros.json"
if not os.path.exists(MACROS_FILE):
    with open(MACROS_FILE, "w") as f:
        json.dump({"grid": {"columns": 6, "rows": 2}, "macros": []}, f, indent=2)
MACROS_UPLOAD_FOLDER = os.path.join(os.getcwd(), "macro-icons")
os.makedirs(MACROS_UPLOAD_FOLDER, exist_ok=True)

# Path to TcNo Account Switcher
switcher_path = r"C:\Program Files\TcNo Account Switcher\TcNo-Acc-Switcher.exe"


macros_lock = Lock()
icon_cache = {}


def get_icon_from_exe(exe_path):
    if exe_path in icon_cache:
        return icon_cache[exe_path]

    try:
        large_icons, _ = win32gui.ExtractIconEx(exe_path, 0)
        if not large_icons:
            return None

        hicon = large_icons[0]
        icon_info = win32gui.GetIconInfo(hicon)
        bmp = win32ui.CreateBitmapFromHandle(icon_info[4])
        bmpinfo = bmp.GetInfo()
        bmpstr = bmp.GetBitmapBits(True)

        img = Image.frombuffer(
            "RGBA",
            (bmpinfo["bmWidth"], bmpinfo["bmHeight"]),
            bmpstr,
            "raw",
            "BGRA",
            0,
            1,
        )
        img = img.transpose(Image.FLIP_TOP_BOTTOM)
        buffered = io.BytesIO()
        img.save(buffered, format="PNG")
        icon_base64 = base64.b64encode(buffered.getvalue()).decode("utf-8")

        icon_cache[exe_path] = icon_base64
        return icon_base64
    except Exception as e:
        print(f"Failed to get high-res icon for {exe_path}: {e}")
        return None


# Launch Applications with a Given Path
@app.route("/open_app", methods=["POST"])
def open_app():
    data = request.get_json(force=True) or {}
    app_path = data.get("app_path", "")
    launch_params = data.get("launch_params", [])

    if not app_path:
        return jsonify({"error": "No application path provided"}), 400

    try:
        if not isinstance(app_path, str):
            return jsonify({"error": "app_path must be a string"}), 400

        app_path = app_path.strip()
        # Be tolerant of clients that accidentally include surrounding quotes.
        app_path = app_path.strip("'\"").strip()

        if not app_path:
            return jsonify({"error": "No application path provided"}), 400

        if isinstance(launch_params, str):
            try:
                params = shlex.split(launch_params, posix=False)
            except ValueError as exc:
                return jsonify({"error": f"Invalid launch_params: {exc}"}), 400
        elif isinstance(launch_params, list):
            params = [str(p) for p in launch_params]
        else:
            return jsonify({"error": "launch_params must be a string or list"}), 400

        is_drive_path = (
            len(app_path) >= 3
            and app_path[1] == ":"
            and app_path[0].isalpha()
            and app_path[2] in ("\\", "/")
        )
        is_uri_target = bool(re.match(r"^[a-zA-Z][a-zA-Z0-9+.-]*:", app_path)) and (
            not is_drive_path
        )
        is_shortcut = app_path.lower().endswith(".lnk")

        if is_uri_target or is_shortcut:
            if params:
                return (
                    jsonify(
                        {
                            "error": "launch_params are not supported for URI or .lnk targets. Use a direct executable path when passing launch parameters."
                        }
                    ),
                    400,
                )
            os.startfile(app_path)
        else:
            cmd = [app_path, *params]
            subprocess.Popen(cmd)

        return (
            jsonify(
                {
                    "status": "success",
                    "message": f"Launched {app_path}",
                    "app_path": app_path,
                    "launch_params": params,
                }
            ),
            200,
        )
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


# Switch Steam User (Auto-login to a specific account)
@app.route("/switch_account", methods=["POST"])
def switch_account():
    steam_id = request.json.get("steam_id")
    if not steam_id:
        return jsonify({"error": "Steam ID is required"}), 400

    try:
        result = subprocess.run(
            [switcher_path, f"+s:{steam_id}"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )

        print("Return Code:", result.returncode)
        print("STDOUT:", result.stdout)
        print("STDERR:", result.stderr)

        if result.stderr.strip() == "":
            return jsonify({"message": "Account switched successfully"}), 200
        else:
            return jsonify(
                {"error": result.stderr.strip(), "code": result.returncode}
            ), 500

    except Exception as e:
        return jsonify({"error": str(e)}), 500


# Macros grid and macros with label, macro, icon
@app.route("/macros", methods=["GET"])
def get_macros():
    with macros_lock:
        try:
            with open(MACROS_FILE, "r") as f:
                data = json.load(f)
            return jsonify(data)
        except Exception as e:
            return jsonify({"error": f"Failed to load macros: {str(e)}"}), 500


# Resize grid
@app.route("/resize_grid", methods=["POST"])
def resize_grid():
    data = request.get_json()
    columns = data.get("columns")
    rows = data.get("rows")

    if (
        not isinstance(columns, int)
        or not isinstance(rows, int)
        or columns < 1
        or rows < 1
    ):
        return jsonify({"error": "Invalid grid dimensions"}), 400

    with macros_lock:
        with open(MACROS_FILE, "r") as f:
            config = json.load(f)

        max_index = columns * rows
        macros = config.get("macros", [])

        if any(m.get("position", 0) >= max_index for m in macros):
            return jsonify(
                {
                    "error": "Grid size too small. Move or delete macros outside bounds first."
                }
            ), 400

        config["grid"] = {"columns": columns, "rows": rows}

        with open(MACROS_FILE, "w") as f:
            json.dump(config, f, indent=2)

    return jsonify({"message": "Grid resized successfully"}), 200


# Create macro with label, macro, icon and position
@app.route("/macros", methods=["POST"])
def add_macro():
    new_macro = request.json
    if not new_macro:
        return jsonify({"error": "No macro data provided"}), 400

    required_keys = ["label", "macro", "icon", "position"]
    if not all(k in new_macro for k in required_keys):
        return jsonify({"error": "Invalid macro format"}), 400

    with macros_lock:
        try:
            with open(MACROS_FILE, "r") as f:
                data = json.load(f)

            # Ensure macros and grid exist
            macros = data.setdefault("macros", [])
            grid = data.setdefault("grid", {"columns": 6, "rows": 2})
            max_slots = grid["columns"] * grid["rows"]

            position = new_macro["position"]

            # Validate position range
            if not isinstance(position, int) or position < 0 or position >= max_slots:
                return jsonify(
                    {"error": f"Position must be between 0 and {max_slots - 1}"}
                ), 400

            # Check for occupied position
            if any(m.get("position") == position for m in macros):
                return jsonify({"error": f"Position {position} is already taken"}), 400

            macros.append(new_macro)

            with open(MACROS_FILE, "w") as f:
                json.dump(data, f, indent=2)

            return jsonify({"message": "Macro added successfully"}), 200

        except Exception as e:
            return jsonify({"error": f"Failed to add macro: {str(e)}"}), 500


# Send icon for macro
@app.route("/macro-icons/<filename>")
def serve_macro_icon(filename):
    return send_from_directory(MACROS_UPLOAD_FOLDER, filename)


# Store icon for macro
@app.route("/upload_macro_icon", methods=["POST"])
def upload_macro_icon():
    if "icon" not in request.files:
        return jsonify({"error": "No file part"}), 400

    file = request.files["icon"]
    if file.filename == "":
        return jsonify({"error": "No selected file"}), 400

    filename = file.filename
    file_path = os.path.join(MACROS_UPLOAD_FOLDER, filename)
    file.save(file_path)

    return jsonify({"icon_path": f"/macro-icons/{filename}"}), 200


# Delete macro
@app.route("/delete_macro", methods=["POST"])
def delete_macro():
    data = request.get_json()
    position = data.get("position")

    if position is None:
        return jsonify({"error": "Position is required"}), 400

    with macros_lock:
        try:
            with open(MACROS_FILE, "r") as f:
                config = json.load(f)

            macros = config.get("macros", [])
            updated_macros = [m for m in macros if m.get("position") != position]
            config["macros"] = updated_macros

            with open(MACROS_FILE, "w") as f:
                json.dump(config, f, indent=2)

            return jsonify({"message": f"Macro at position {position} deleted"}), 200
        except Exception as e:
            return jsonify({"error": str(e)}), 500


# Swap position of two macros
@app.route("/swap_macros", methods=["POST"])
def swap_macros():
    data = request.get_json()
    from_pos = data.get("from")
    to_pos = data.get("to")

    if from_pos is None or to_pos is None:
        return jsonify({"error": "Invalid positions"}), 400

    with macros_lock:
        try:
            with open(MACROS_FILE, "r") as f:
                config = json.load(f)

            macros = config.get("macros", [])

            macro_from = next(
                (m for m in macros if m.get("position") == from_pos), None
            )
            macro_to = next((m for m in macros if m.get("position") == to_pos), None)

            if macro_from:
                macro_from["position"] = to_pos
            if macro_to:
                macro_to["position"] = from_pos

            with open(MACROS_FILE, "w") as f:
                json.dump(config, f, indent=2)

            return jsonify({"message": "Swap complete"}), 200
        except Exception as e:
            return jsonify({"error": str(e)}), 500


# Audio session name, PID, cached icon, volume
@app.route("/audio_sessions_metadata", methods=["GET"])
def get_audio_sessions_metadata():
    comtypes.CoInitialize()

    sessions = AudioUtilities.GetAllSessions()
    results = []

    for session in sessions:
        if session.Process:
            try:
                volume = session._ctl.QueryInterface(
                    ISimpleAudioVolume
                ).GetMasterVolume()
                exe_path = session.Process.exe()
                icon = get_icon_from_exe(exe_path)
                results.append(
                    {
                        "name": session.Process.name(),
                        "pid": session.Process.pid,
                        "icon": icon,
                        "volume": round(volume * 100, 2),
                    }
                )
            except Exception as e:
                print(f"Error in metadata fetch: {e}")

    return jsonify(results)


# Session name and volume only
@app.route("/audio_sessions_volume", methods=["GET"])
def get_audio_sessions_volume():
    comtypes.CoInitialize()

    sessions = AudioUtilities.GetAllSessions()
    results = []

    for session in sessions:
        if session.Process:
            try:
                volume = session._ctl.QueryInterface(
                    ISimpleAudioVolume
                ).GetMasterVolume()
                results.append(
                    {"name": session.Process.name(), "volume": round(volume * 100, 2)}
                )
            except Exception as e:
                print(f"Error in volume fetch: {e}")

    return jsonify(results)


# Set volume for a specific app
@app.route("/set_app_volume", methods=["POST"])
def set_app_volume():
    data = request.json
    app_name = data.get("app_name")
    volume_level = float(data.get("volume")) / 100.0

    sessions = AudioUtilities.GetAllSessions()

    for session in sessions:
        if session.Process and session.Process.name().lower() == app_name.lower():
            volume = session._ctl.QueryInterface(ISimpleAudioVolume)
            volume.SetMasterVolume(volume_level, None)
            return jsonify(
                {"message": f"Volume set for {app_name} to {volume_level * 100}%"}
            )

    return jsonify({"error": f"{app_name} not found"}), 404


# Set master volume
@app.route("/set_master_volume", methods=["POST"])
def set_master_volume():
    comtypes.CoInitialize()

    volume_level = float(request.json.get("volume")) / 100.0

    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))

    volume.SetMasterVolumeLevelScalar(volume_level, None)

    return jsonify({"message": f"Master volume set to {volume_level * 100}%"})


# Get audio outputs
@app.route("/audio_output_devices", methods=["GET"])
def list_playback_devices():
    comtypes.CoInitialize()
    try:
        # Default output device (multimedia render)
        default_device = AudioUtilities.GetSpeakers()
        default_id = default_device.id

        # Active render (playback) devices
        with warnings.catch_warnings():
            warnings.simplefilter(
                "ignore", UserWarning
            )  # pycaw may warn on odd devices
            devices = AudioUtilities.GetAllDevices(
                data_flow=EDataFlow.eRender.value,
                device_state=DEVICE_STATE.ACTIVE.value,
            )

        payload = [
            {"id": d.id, "name": d.FriendlyName, "is_default": d.id == default_id}
            for d in devices
        ]
        return jsonify(payload), 200
    finally:
        comtypes.CoUninitialize()


# Set audio output
@app.route("/set_audio_output_device", methods=["POST"])
def set_playback_device():
    data = request.get_json(force=True) or {}
    dev_id = data.get("device_id")
    if not dev_id:
        return jsonify({"error": "device_id missing"}), 400

    comtypes.CoInitialize()
    try:
        AudioUtilities.SetDefaultDevice(dev_id)
        return jsonify({"message": f"default output set to {dev_id}"}), 200
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500
    finally:
        comtypes.CoUninitialize()


# Media keys
def press_media_key(vk_code):
    ctypes.windll.user32.keybd_event(vk_code, 0, 0, 0)  # key down
    ctypes.windll.user32.keybd_event(vk_code, 0, 2, 0)  # key up


@app.route("/media/play_pause", methods=["POST"])
def media_play_pause():
    press_media_key(0xB3)
    return jsonify({"message": "Play/Pause toggled"}), 200


@app.route("/media/next", methods=["POST"])
def media_next():
    press_media_key(0xB0)
    return jsonify({"message": "Next track"}), 200


@app.route("/media/prev", methods=["POST"])
def media_prev():
    press_media_key(0xB1)
    return jsonify({"message": "Previous track"}), 200


# Simulate Keyboard Key Press
@app.route("/press_key", methods=["POST"])
def press_key():
    data = request.json
    key = data.get("key")

    if key:
        try:
            pyautogui.press(key)
            return jsonify({"status": "success", "message": f"Pressed {key}"}), 200
        except Exception as e:
            return jsonify({"error": str(e)}), 500
    else:
        return jsonify({"error": "No key provided"}), 400


# Run Custom Commands
@app.route("/run_command", methods=["POST"])
def run_command():
    data = request.json
    command = data.get("command")

    if command:
        try:
            subprocess.Popen(command, shell=True)
            return jsonify(
                {"status": "success", "message": f"Executed: {command}"}
            ), 200
        except Exception as e:
            return jsonify({"error": str(e)}), 500
    else:
        return jsonify({"error": "No command provided"}), 400


@app.route("/type_text", methods=["POST"])
def type_text():
    data = request.get_json(force=True) or {}
    text = data.get("text", "")

    if not isinstance(text, str) or text == "":
        return jsonify({"error": "text is required"}), 400

    try:
        parts = COMMAND_PATTERN.split(text)

        for part in parts:
            if part == "enter":
                keyboard.press_and_release("enter")

            elif part.startswith("wait:"):
                ms = int(part.split(":")[1])
                time.sleep(ms / 1000)

            else:
                keyboard.write(part, delay=0.01)

        return jsonify({"message": "Macro executed"}), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001, debug=False)
