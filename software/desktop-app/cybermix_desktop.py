"""
CyberMix Desktop Controller
- Controls Windows application volumes
- Interfaces with CyberMix firmware via USB serial
- Updates OLED screens on the board
"""

import serial
import time
import threading
import sys
import struct
import win32api
import win32con
import win32gui
import win32process
import win32com.client
import ctypes

# --- USB Serial Port ---
SERIAL_PORT = 'COM3'  # Change to your board's port
BAUDRATE = 115200

# --- MIDI CC Mapping ---
FADER_CC = [20, 21, 22, 23, 24]

# --- Windows Audio Control ---
# Uses pycaw for per-app volume control
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, ISimpleAudioVolume
from comtypes import CLSCTX_ALL

def set_app_volume(app_name, volume):
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        if session.Process and app_name.lower() in session.Process.name().lower():
            volume_obj = session._ctl.QueryInterface(ISimpleAudioVolume)
            volume_obj.SetMasterVolume(float(volume), None)
            return True
    return False

def get_app_volumes():
    sessions = AudioUtilities.GetAllSessions()
    app_vols = {}
    for session in sessions:
        if session.Process:
            name = session.Process.name()
            volume_obj = session._ctl.QueryInterface(ISimpleAudioVolume)
            app_vols[name] = volume_obj.GetMasterVolume()
    return app_vols

# --- Serial Communication ---
def send_screen_update(ser, screen_idx, text):
    if ser:
        msg = b'S' + bytes([screen_idx]) + text.encode('utf-8')[:16]
        ser.write(msg)

def read_fader_data(ser):
    if ser and ser.in_waiting >= 6:
        data = ser.read(6)
        if data[0] == ord('F'):
            return list(data[1:])
    return None

def find_serial_port():
    import serial.tools.list_ports
    ports = list(serial.tools.list_ports.comports())
    for p in ports:
        if 'XIAO' in p.description or 'USB Serial' in p.description:
            return p.device
    return None

def main():
    ser = None
    app_names = ['chrome.exe', 'discord.exe', 'spotify.exe', 'vlc.exe', 'system']
    last_vols = [0.0] * 5
    print('CyberMix Desktop Controller started. Running in test mode until device is found.')
    while True:
        if ser is None:
            port = find_serial_port()
            if port:
                try:
                    ser = serial.Serial(port, BAUDRATE, timeout=0.1)
                    print(f'Connected to CyberMix device on {port}')
                except Exception as e:
                    print(f'Failed to connect: {e}')
                    ser = None
            else:
                # Simulate fader values in test mode
                fader_vals = [int(127 * (i+1)/5) for i in range(5)]
        else:
            fader_vals = read_fader_data(ser)
            if fader_vals is None:
                # If no data, simulate
                fader_vals = [int(127 * (i+1)/5) for i in range(5)]
        for i, val in enumerate(fader_vals):
            vol = val / 127.0
            set_app_volume(app_names[i], vol)
            last_vols[i] = vol
            send_screen_update(ser, i, f"{app_names[i][:8]}: {int(vol*100)}%")
        time.sleep(0.5)

if __name__ == "__main__":
    main()
