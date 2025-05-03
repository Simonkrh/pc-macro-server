from flask import Flask, request, jsonify
from flask_cors import CORS
import os
import subprocess
import pyautogui
import comtypes
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume, IAudioEndpointVolume
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
import ctypes
import base64
from PIL import Image
import win32gui
import win32ui
import io

app = Flask(__name__)
CORS(app)

# Path to TcNo Account Switcher
switcher_path = r"C:\Program Files\TcNo Account Switcher\TcNo-Acc-Switcher.exe"

def get_icon_from_exe(exe_path):
    try:
        large_icons, small_icons = win32gui.ExtractIconEx(exe_path, 0)
        if not large_icons:
            return None

        hicon = large_icons[0]

        # Get icon info
        icon_info = win32gui.GetIconInfo(hicon)
        bmp = win32ui.CreateBitmapFromHandle(icon_info[4])  # hbmColor

        bmpinfo = bmp.GetInfo()
        bmpstr = bmp.GetBitmapBits(True)

        img = Image.frombuffer(
            'RGBA',
            (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
            bmpstr, 'raw', 'BGRA', 0, 1
        )
        img = img.transpose(Image.FLIP_TOP_BOTTOM)

        # Save to base64
        buffered = io.BytesIO()
        img.save(buffered, format="PNG")
        icon_base64 = base64.b64encode(buffered.getvalue()).decode('utf-8')

        return icon_base64
    except Exception as e:
        print(f"Failed to get high-res icon: {e}")
        return None

    
# Launch Applications with a Given Path
@app.route('/open_app', methods=['POST'])
def open_app():
    data = request.json
    app_path = data.get('app_path')

    if not app_path:
        return jsonify({"error": "No application path provided"}), 400

    try:
        os.startfile(app_path)
        return jsonify({"status": "success", "message": f"Launched {app_path}"}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# Switch Steam User (Auto-login to a specific account)
@app.route('/switch_account', methods=['POST'])
def switch_account():
    steam_id = request.json.get('steam_id')
    if not steam_id:
        return jsonify({'error': 'Steam ID is required'}), 400

    try:
        result = subprocess.run(
            [switcher_path, f"+s:{steam_id}"],
            capture_output=True,
            text=True
        )

        print("Return Code:", result.returncode)
        print("STDOUT:", result.stdout)
        print("STDERR:", result.stderr)

        if result.stderr.strip() == "":
            return jsonify({'message': 'Account switched successfully'}), 200
        else:
            return jsonify({'error': result.stderr.strip(), 'code': result.returncode}), 500

    except Exception as e:
        return jsonify({'error': str(e)}), 500


# Simulate Keyboard Key Press
@app.route('/press_key', methods=['POST'])
def press_key():
    data = request.json
    key = data.get('key')

    if key:
        try:
            pyautogui.press(key)
            return jsonify({"status": "success", "message": f"Pressed {key}"}), 200
        except Exception as e:
            return jsonify({"error": str(e)}), 500
    else:
        return jsonify({"error": "No key provided"}), 400

# Run Custom Commands 
@app.route('/run_command', methods=['POST'])
def run_command():
    data = request.json
    command = data.get('command')

    if command:
        try:
            subprocess.Popen(command, shell=True)
            return jsonify({"status": "success", "message": f"Executed: {command}"}), 200
        except Exception as e:
            return jsonify({"error": str(e)}), 500
    else:
        return jsonify({"error": "No command provided"}), 400

# Get list of current audio sessions (apps)
@app.route('/audio_sessions', methods=['GET'])
def get_audio_sessions():
    comtypes.CoInitialize()

    sessions = AudioUtilities.GetAllSessions()
    results = []

    for session in sessions:
        if session.Process:
            try:
                exe_path = session.Process.exe()
                icon = get_icon_from_exe(exe_path)
            except:
                icon = None

            volume = session._ctl.QueryInterface(ISimpleAudioVolume).GetMasterVolume()
            results.append({
                "name": session.Process.name(),
                "pid": session.Process.pid,
                "volume": round(volume * 100, 2),
                "icon": icon
            })
    
    return jsonify(results)

# Set volume for a specific app
@app.route('/set_app_volume', methods=['POST'])
def set_app_volume():
    data = request.json
    app_name = data.get("app_name")
    volume_level = float(data.get("volume")) / 100.0 

    sessions = AudioUtilities.GetAllSessions()

    for session in sessions:
        if session.Process and session.Process.name().lower() == app_name.lower():
            volume = session._ctl.QueryInterface(ISimpleAudioVolume)
            volume.SetMasterVolume(volume_level, None)
            return jsonify({"message": f"Volume set for {app_name} to {volume_level * 100}%"})

    return jsonify({"error": f"{app_name} not found"}), 404

# Set master volume
@app.route('/set_master_volume', methods=['POST'])
def set_master_volume():
    comtypes.CoInitialize() 

    volume_level = float(request.json.get("volume")) / 100.0  

    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))

    volume.SetMasterVolumeLevelScalar(volume_level, None)

    return jsonify({"message": f"Master volume set to {volume_level * 100}%"})

# Media keys
def press_media_key(vk_code):
    ctypes.windll.user32.keybd_event(vk_code, 0, 0, 0) # key down
    ctypes.windll.user32.keybd_event(vk_code, 0, 2, 0) # key up

@app.route('/media/play_pause', methods=['POST'])
def media_play_pause():
    press_media_key(0xB3)
    return jsonify({"message": "Play/Pause toggled"}), 200

@app.route('/media/next', methods=['POST'])
def media_next():
    press_media_key(0xB0)
    return jsonify({"message": "Next track"}), 200

@app.route('/media/prev', methods=['POST'])
def media_prev():
    press_media_key(0xB1)
    return jsonify({"message": "Previous track"}), 200

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=5001, debug=True)
