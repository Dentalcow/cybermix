"""
CyberMix Desktop Controller
- Combined GUI and CLI controller for Windows application volumes
- Interfaces with CyberMix firmware via USB serial
- Real-time fader sliders with cyberpunk UI or headless CLI mode
- Updates OLED screens on the board
"""

import sys
import threading
import time
import json
import os
import argparse
import serial
import serial.tools.list_ports

try:
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, ISimpleAudioVolume
    from comtypes import CLSCTX_ALL
except ImportError:
    AudioUtilities = None
    ISimpleAudioVolume = None

import psutil

# --- Configuration ---
BAUDRATE = 115200
APP_NAMES = ['chrome.exe', 'discord.exe', 'spotify.exe', 'vlc.exe', 'system']
FADER_CC = [20, 21, 22, 23, 24]
FADER_COUNT = 5
SETTINGS_FILE = os.path.join(os.path.dirname(__file__), 'cybermix_settings.json')


# --- Serial Communication Functions ---
def find_serial_port():
    """Find CyberMix device on serial ports"""
    ports = list(serial.tools.list_ports.comports())
    for p in ports:
        if 'XIAO' in p.description or 'USB Serial' in p.description:
            return p.device
    return None


def send_screen_update(ser, screen_idx, text):
    """Send text update to OLED screen"""
    if ser:
        msg = b'S' + bytes([screen_idx]) + text.encode('utf-8')[:16]
        ser.write(msg)


def read_fader_data(ser):
    """Read fader data from serial port"""
    if ser and ser.in_waiting >= 6:
        data = ser.read(6)
        if data[0] == ord('F'):
            return list(data[1:])
    return None


# --- Windows Audio Control Functions ---
def set_app_volume(app_name, volume):
    """Set volume for specific application"""
    if not AudioUtilities:
        return False
    
    if app_name == 'System Master Volume':
        try:
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            import comtypes
            master = comtypes.cast(interface, comtypes.POINTER(IAudioEndpointVolume))
            master.SetMasterVolumeLevelScalar(float(volume), None)
            return True
        except:
            return False
    
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        if session.Process and app_name.lower() in session.Process.name().lower():
            try:
                volume_obj = session._ctl.QueryInterface(ISimpleAudioVolume)
                volume_obj.SetMasterVolume(float(volume), None)
                return True
            except:
                pass
    return False


def get_app_volumes():
    """Get current volumes for all applications"""
    if not AudioUtilities:
        return {}
    
    app_vols = {}
    try:
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            if session.Process:
                name = session.Process.name()
                try:
                    volume_obj = session._ctl.QueryInterface(ISimpleAudioVolume)
                    app_vols[name] = volume_obj.GetMasterVolume()
                except:
                    pass
    except:
        pass
    return app_vols


# --- Headless CLI Controller ---
class DesktopController:
    """Headless controller for CLI mode"""
    def __init__(self):
        self.ser = None
        self.app_names = APP_NAMES
        self.running = True
        self.fader_vals = [64] * 5

    def connect(self):
        """Connect to serial device"""
        port = find_serial_port()
        if port:
            try:
                self.ser = serial.Serial(port, BAUDRATE, timeout=0.1)
                print(f'✓ Connected to CyberMix device on {port}')
                return True
            except Exception as e:
                print(f'✗ Failed to connect: {e}')
                self.ser = None
                return False
        else:
            print('✗ CyberMix device not found on any serial port')
            return False

    def run(self):
        """Main control loop"""
        print('CyberMix Desktop Controller started')
        print('Running in CLI mode (use --gui for GUI)')
        print('')
        
        attempt = 0
        while self.running:
            if self.ser is None:
                if attempt % 20 == 0:  # Try to connect every 10 seconds
                    self.connect()
                attempt += 1
            else:
                fader_vals = read_fader_data(self.ser)
                if fader_vals is None:
                    fader_vals = [64] * 5
                
                for i, val in enumerate(fader_vals):
                    vol = val / 127.0
                    set_app_volume(self.app_names[i], vol)
                    send_screen_update(self.ser, i, f"{self.app_names[i][:8]}: {int(vol*100)}%")
                
                self.fader_vals = fader_vals

            time.sleep(0.5)

    def stop(self):
        """Stop controller"""
        self.running = False
        if self.ser:
            self.ser.close()


# --- GUI Mode ---
def run_gui():
    """Launch GUI version"""
    try:
        from PyQt5 import QtWidgets, QtCore, QtGui
        from PyQt5.QtWidgets import QSystemTrayIcon
        from PyQt5.QtGui import QIcon
    except ImportError:
        print('PyQt5 not installed. Install with: pip install PyQt5')
        return

    class SerialWorker(QtCore.QObject):
        fader_update = QtCore.pyqtSignal(list)
        connection_status = QtCore.pyqtSignal(bool)
        
        def __init__(self):
            super().__init__()
            self.ser = None
            self.running = True

        def find_serial_port(self):
            return find_serial_port()

        def run(self):
            while self.running:
                if self.ser is None:
                    port = self.find_serial_port()
                    if port:
                        try:
                            self.ser = serial.Serial(port, BAUDRATE, timeout=0.1)
                            self.connection_status.emit(True)
                        except:
                            self.ser = None
                            self.connection_status.emit(False)
                    else:
                        self.connection_status.emit(False)
                        time.sleep(1)
                else:
                    if self.ser.in_waiting >= 6:
                        data = self.ser.read(6)
                        if data[0] == ord('F'):
                            vals = list(data[1:])
                            self.fader_update.emit(vals)
                    time.sleep(0.05)

        def stop(self):
            self.running = False
            if self.ser:
                self.ser.close()

    class DraggableLabel(QtWidgets.QLabel):
        def __init__(self, text, parent=None):
            super().__init__(text, parent)
            self.setAcceptDrops(True)
            self.setAlignment(QtCore.Qt.AlignCenter)
            self.setStyleSheet('background: #222; color: #e0e0ff; border: 2px dashed #ff2e63;')

        def dragEnterEvent(self, event):
            if event.mimeData().hasText():
                event.acceptProposedAction()
                self.setStyleSheet('background: #222; color: #08d9d6; border: 2px solid #08d9d6;')
            else:
                event.ignore()

        def dragMoveEvent(self, event):
            if event.mimeData().hasText():
                event.acceptProposedAction()
            else:
                event.ignore()

        def dropEvent(self, event):
            if event.mimeData().hasText():
                self.setText(event.mimeData().text())
                self.setStyleSheet('background: #222; color: #e0e0ff; border: 2px dashed #ff2e63;')
                self.parent().process_dropped(self, event.mimeData().text())
                event.acceptProposedAction()
            else:
                event.ignore()

        def dragLeaveEvent(self, event):
            self.setStyleSheet('background: #222; color: #e0e0ff; border: 2px dashed #ff2e63;')

    class ProcessListWidget(QtWidgets.QListWidget):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self.setDragEnabled(True)

        def startDrag(self, supportedActions):
            item = self.currentItem()
            if item:
                mime = QtCore.QMimeData()
                mime.setText(item.text())
                drag = QtGui.QDrag(self)
                drag.setMimeData(mime)
                drag.exec_(QtCore.Qt.CopyAction)

    class MainWindow(QtWidgets.QWidget):
        def __init__(self):
            super().__init__()
            self.setWindowTitle('CyberMix Volume Mixer')
            self.setStyleSheet('background-color: #18181c; color: #e0e0ff; font-family: "Orbitron", "Segoe UI", sans-serif;')
            self.resize(700, 500)
            self.page = 0
            self.processes = self.get_processes()
            self.fader_assignments = ['System Master Volume'] + [None] * (FADER_COUNT - 1)
            
            # Setup system tray
            self.tray_icon = QSystemTrayIcon(self)
            
            # Create a simple icon (colored pixel)
            pixmap = QtGui.QPixmap(16, 16)
            pixmap.fill(QtGui.QColor(8, 217, 214))  # Cyan
            icon = QIcon(pixmap)
            self.tray_icon.setIcon(icon)
            
            tray_menu = QtWidgets.QMenu()
            show_action = tray_menu.addAction('Show')
            show_action.triggered.connect(self.show_window)
            hide_action = tray_menu.addAction('Hide')
            hide_action.triggered.connect(self.hide_to_tray)
            tray_menu.addSeparator()
            exit_action = tray_menu.addAction('Exit')
            exit_action.triggered.connect(self.exit_app)
            self.tray_icon.setContextMenu(tray_menu)
            self.tray_icon.show()
            
            main_layout = QtWidgets.QHBoxLayout(self)
            
            # Side menu: process list
            self.process_list = ProcessListWidget()
            self.process_list.setFixedWidth(180)
            self.process_list.addItems(self.get_processes())
            self.process_list.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
            main_layout.addWidget(self.process_list)
            
            # Fader area
            fader_area = QtWidgets.QVBoxLayout()
            self.status_label = QtWidgets.QLabel('Device: Not Connected')
            self.status_label.setStyleSheet('color: #ff2e63; font-weight: bold;')
            fader_area.addWidget(self.status_label)
            
            fader_layout = QtWidgets.QHBoxLayout()
            self.fader_sliders = []
            self.fader_labels = []
            self.oled_previews = []
            
            for i in range(FADER_COUNT):
                vbox = QtWidgets.QVBoxLayout()
                label = DraggableLabel('System Master Volume' if i == 0 else '---', self)
                vbox.addWidget(label)
                self.fader_labels.append(label)
                
                slider = QtWidgets.QSlider(QtCore.Qt.Vertical)
                slider.setMinimum(0)
                slider.setMaximum(127)
                slider.setValue(64)
                slider.setStyleSheet('QSlider::groove:vertical {background: #222; border: 1px solid #ff2e63;} QSlider::handle:vertical {background: #08d9d6; border: 1px solid #ff2e63;}')
                slider.valueChanged.connect(lambda val, idx=i: self.set_app_volume(idx, val))
                vbox.addWidget(slider)
                self.fader_sliders.append(slider)
                
                preview = QtWidgets.QLabel('OLED Preview')
                preview.setAlignment(QtCore.Qt.AlignCenter)
                preview.setStyleSheet('background: #222; color: #08d9d6; border: 1px solid #ff2e63;')
                vbox.addWidget(preview)
                self.oled_previews.append(preview)
                
                fader_layout.addLayout(vbox)
            
            fader_area.addLayout(fader_layout)
            
            nav_layout = QtWidgets.QHBoxLayout()
            self.prev_btn = QtWidgets.QPushButton('Prev')
            self.prev_btn.clicked.connect(self.prev_page)
            self.next_btn = QtWidgets.QPushButton('Next')
            self.next_btn.clicked.connect(self.next_page)
            nav_layout.addWidget(self.prev_btn)
            nav_layout.addWidget(self.next_btn)
            fader_area.addLayout(nav_layout)
            
            # LED bar at bottom
            self.led_bar = QtWidgets.QProgressBar()
            self.led_bar.setMaximum(100)
            self.led_bar.setTextVisible(False)
            self.led_bar.setStyleSheet('QProgressBar {background: #222; border: 1px solid #ff2e63;} QProgressBar::chunk {background: #08d9d6;}')
            fader_area.addWidget(self.led_bar)
            
            # LED simulation widget
            self.led_sim = QtWidgets.QWidget()
            self.led_sim.setFixedHeight(40)
            self.led_sim.paintEvent = lambda event: self.paint_led_sim(self.led_sim, self.led_bar.value())
            fader_area.addWidget(self.led_sim)
            
            main_layout.addLayout(fader_area)
            self.setLayout(main_layout)
            
            self.serial_worker = SerialWorker()
            self.thread = QtCore.QThread()
            self.serial_worker.moveToThread(self.thread)
            self.serial_worker.fader_update.connect(self.update_faders)
            self.serial_worker.connection_status.connect(self.update_status)
            self.thread.started.connect(self.serial_worker.run)
            self.thread.start()
            
            self.load_settings()

        def is_user_app(self, proc_name):
            """Check if process is a user-facing app (like Task Manager basic view)"""
            # Exclude system/background processes
            exclude_keywords = [
                'system', 'svchost', 'wininit', 'winlogon', 'lsass', 'csrss',
                'smss', 'services', 'ntoskrnl', 'conhost', 'dwm', 'explorer',
                'taskhostw', 'rundll32', 'dllhost', 'wmiprvse', 'sihost',
                'wuauserv', 'spoolsv', 'searchindexer', 'vds', 'nvda',
                'nvidia', 'amd', 'intel', 'realtek', 'qualcomm', 'razer',
                'corsair', 'msi', 'asus', 'lenovo', 'dell', 'hp', 'acer',
                'microsoft edge', 'windowsapps', 'host', 'service', 'daemon',
                'agent', 'monitor', 'update', 'defender', 'antimalware',
                'bits', 'trustedinstaller', 'ngscryptoapi'
            ]
            
            lower_name = proc_name.lower()
            
            # Exclude known system processes
            for keyword in exclude_keywords:
                if keyword in lower_name:
                    return False
            
            # Include apps with known user-facing extensions
            if lower_name.endswith(('.exe', '.EXE')):
                # Further filter: look for common user apps or apps with audio sessions
                user_apps = [
                    'chrome', 'firefox', 'edge', 'safari', 'opera', 'brave',
                    'discord', 'slack', 'teams', 'zoom', 'skype', 'telegram',
                    'spotify', 'youtube', 'vlc', 'obs', 'audacity', 'steam',
                    'epic', 'uplay', 'battlenet', 'blender', 'adobe', 'gimp',
                    'logic', 'ableton', 'cubase', 'reaper', 'fl studio',
                    'finale', 'notion', 'slack', 'messenger', 'whatsapp',
                    'handbrake', 'ffmpeg', 'potplayer', 'winamp', 'foobar',
                    'mediamonkey', 'musicbee', 'itunes', 'groove', 'amazon music',
                    'twitch', 'obs', 'xsplit', 'streamlabs', 'dexpot', 'rainmeter'
                ]
                
                # Check if it's a known user app
                for app in user_apps:
                    if app in lower_name:
                        return True
                
                # If it has an audio session, it's probably worth showing
                # (This gets checked separately)
                return None  # Let audio session check decide
            
            return False

        def get_processes(self):
            names = ['System Master Volume']
            seen = set()
            audio_names = []
            unknown_apps = []
            
            if AudioUtilities:
                try:
                    sessions = AudioUtilities.GetAllSessions()
                    for session in sessions:
                        if session.Process:
                            pname = session.Process.name()
                            if pname and pname.lower() not in seen:
                                # If app has audio session, it's worth including
                                if self.is_user_app(pname) is not False:
                                    audio_names.append(pname)
                                    seen.add(pname.lower())
                except:
                    pass
            
            try:
                for proc in psutil.process_iter(['name']):
                    pname = proc.info['name']
                    if pname and pname.lower() not in seen:
                        if self.is_user_app(pname) is True:
                            unknown_apps.append(pname)
                            seen.add(pname.lower())
            except:
                pass
            
            return ['System Master Volume'] + sorted(audio_names, key=str.lower) + sorted(unknown_apps, key=str.lower)

        def set_app_volume(self, idx, val):
            proc = self.fader_assignments[idx]
            if proc:
                set_app_volume(proc, val / 127.0)
            self.oled_previews[idx].setText(f'{proc[:8] if proc else "---"}: {int(val/127.0*100)}%')
            self.save_settings()
            self.update_led_bar()

        def update_faders(self, vals):
            for i, val in enumerate(vals):
                self.fader_sliders[i].blockSignals(True)
                self.fader_sliders[i].setValue(val)
                self.fader_sliders[i].blockSignals(False)
                proc = self.fader_assignments[i]
                self.oled_previews[i].setText(f'{proc[:8] if proc else "---"}: {int(val/127.0*100)}%')
            self.update_led_bar()

        def update_led_bar(self):
            if hasattr(self, 'led_bar'):
                val = self.fader_sliders[0].value()
                percent = int(val / 127.0 * 100)
                self.led_bar.setValue(percent)

        def closeEvent(self, event):
            if self.tray_icon.isVisible():
                self.hide_to_tray()
                event.ignore()
            else:
                self.serial_worker.stop()
                self.thread.quit()
                self.thread.wait()
                event.accept()
        
        def changeEvent(self, event):
            if event.type() == QtCore.QEvent.WindowStateChange:
                if self.isMinimized():
                    self.hide_to_tray()
                    event.ignore()
        
        def hide_to_tray(self):
            self.hide()
            self.setWindowState(QtCore.Qt.WindowNoState)
        
        def show_window(self):
            self.showNormal()
            self.activateWindow()
        
        def exit_app(self):
            self.serial_worker.stop()
            self.thread.quit()
            self.thread.wait()
            QtWidgets.QApplication.quit()

        def prev_page(self):
            if self.page > 0:
                self.page -= 1
                self.refresh_page()

        def next_page(self):
            max_page = max(0, (len(self.processes) - 1) // 5)
            if self.page < max_page:
                self.page += 1
                self.refresh_page()

        def refresh_page(self):
            procs = self.processes[self.page * 5:(self.page + 1) * 5]
            for i, label in enumerate(self.fader_labels):
                if i < len(procs):
                    if self.fader_assignments[i] in procs:
                        label.setText(self.fader_assignments[i])
                    elif i == 0:
                        self.fader_assignments[i] = procs[0]
                        label.setText(procs[0])
                    else:
                        self.fader_assignments[i] = None
                        label.setText('---')
                else:
                    self.fader_assignments[i] = None
                    label.setText('---')
            
            self.process_list.clear()
            self.process_list.addItems(self.get_processes())

        def save_settings(self):
            data = {
                'fader_assignments': self.fader_assignments,
                'fader_values': [slider.value() for slider in self.fader_sliders],
                'page': self.page
            }
            try:
                with open(SETTINGS_FILE, 'w') as f:
                    json.dump(data, f)
            except:
                pass

        def load_settings(self):
            if os.path.exists(SETTINGS_FILE):
                try:
                    with open(SETTINGS_FILE, 'r') as f:
                        data = json.load(f)
                    self.fader_assignments = data.get('fader_assignments', self.fader_assignments)
                    self.page = data.get('page', self.page)
                    for i, val in enumerate(data.get('fader_values', [])):
                        if i < len(self.fader_sliders):
                            self.fader_sliders[i].setValue(val)
                except:
                    pass

        def update_status(self, connected):
            if connected:
                self.status_label.setText('Device: Connected')
                self.status_label.setStyleSheet('color: #08d9d6; font-weight: bold;')
            else:
                self.status_label.setText('Device: Not Connected')
                self.status_label.setStyleSheet('color: #ff2e63; font-weight: bold;')

        def paint_led_sim(self, widget, percent):
            painter = QtGui.QPainter(widget)
            led_count = 8
            margin = 8
            spacing = 8
            w = widget.width()
            h = widget.height()
            led_w = (w - margin * 2 - spacing * (led_count - 1)) // led_count
            led_h = h - margin * 2
            num_on = int(percent / 100 * led_count)
            for i in range(led_count):
                x = margin + i * (led_w + spacing)
                color = QtGui.QColor(8, 217, 214) if i < num_on else QtGui.QColor(32, 32, 32)
                painter.setBrush(color)
                painter.setPen(QtGui.QColor('#ff2e63'))
                painter.drawRect(x, margin, led_w, led_h)
            painter.end()

        def process_dropped(self, label, proc_name):
            idx = self.fader_labels.index(label)
            self.fader_assignments[idx] = proc_name
            label.setText(proc_name)
            self.save_settings()

    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    
    # Timer to update LED simulation
    timer = QtCore.QTimer()
    timer.timeout.connect(lambda: window.led_sim.update())
    timer.start(50)
    
    window.show()
    sys.exit(app.exec_())


# --- Main Entry Point ---
if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='CyberMix Volume Mixer')
    parser.add_argument('--gui', action='store_true', help='Launch GUI mode (default if PyQt5 available)')
    parser.add_argument('--cli', action='store_true', help='Launch CLI mode (headless)')
    
    args = parser.parse_args()
    
    # Determine mode
    use_gui = args.gui
    if not args.cli and not args.gui:
        # Auto-detect: try GUI first, fall back to CLI
        try:
            import PyQt5
            use_gui = True
        except ImportError:
            use_gui = False
    
    if use_gui:
        run_gui()
    else:
        controller = DesktopController()
        try:
            controller.run()
        except KeyboardInterrupt:
            print('\n✓ CyberMix stopped')
            controller.stop()
