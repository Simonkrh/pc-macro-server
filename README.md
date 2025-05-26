# PC Macro Server

This is a lightweight Flask API for sending macro commands to a PC. It supports launching applications, switching Steam accounts (via [TcNo Account Switcher](https://github.com/TCNOco/TcNo-Acc-Switcher/releases/tag/2024-08-30_01)), controlling per-app volume, sending media keys, and simulating keypresses.

This server is used by my [PC Monitor Dashboard](https://github.com/Simonkrh/pc-monitor-dashboard) project, which runs on a Raspberry Pi touchscreen as a dashboard for my gaming PC.

## How to Use
**1. Set the path to the TcNo Account Switcher**  
Near the top of `macro_server.py`, update this line to match the install location:
```python
switcher_path = r"C:\Program Files\TcNo Account Switcher\TcNo-Acc-Switcher.exe"
```

**2 Install requirments**
```bash
pip install -r requirements.txt
```

**2. Run the Server**

```bash
python macro_server.py
