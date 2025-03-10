from flask import Flask, request, jsonify
from flask_cors import CORS
import os
import time
import re
import subprocess
import pyautogui

app = Flask(__name__)
CORS(app)

# Path to TcNo Account Switcher
switcher_path = r"C:\Program Files\TcNo Account Switcher\TcNo-Acc-Switcher.exe"

# Launch Applications with a Given Path
@app.route('/open_app', methods=['POST'])
def open_app():
    data = request.json
    app_path = data.get('app_path')  # Expecting full path

    if not app_path:
        return jsonify({"error": "No application path provided"}), 400

    try:
        subprocess.Popen(app_path, shell=True)
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
        if result.returncode == 0:
            return jsonify({'message': 'Account switched successfully'}), 200
        else:
            return jsonify({'error': result.stderr.strip()}), 500
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

# Start Flask Server
if __name__ == '__main__':
    app.run(host="0.0.0.0", port=5002, debug=True)
