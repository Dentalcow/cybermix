"""
CyberMix Desktop Controller GUI
- Modern cyberpunk-inspired UI for controlling Windows app volumes
- Interfaces with CyberMix firmware via USB serial
- Real-time fader sliders, OLED screen preview, connection status
"""

import sys
import threading
import time
import json
import os
from PyQt5 import QtWidgets, QtCore, QtGui

try:
    import serial
    import serial.tools.list_ports
except ImportError:
    serial = None

try:
    from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
except ImportError:
    AudioUtilities = None
    ISimpleAudioVolume = None

import psutil

APP_NAMES = ['chrome.exe', 'discord.exe', 'spotify.exe', 'vlc.exe', 'system']
FADER_CC = [20, 21, 22, 23, 24]
BAUDRATE = 115200
SETTINGS_FILE = os.path.join(os.path.dirname(__file__), 'cybermix_settings.json')
FADER_COUNT = 5

class SerialWorker(QtCore.QObject):
    fader_update = QtCore.pyqtSignal(list)
    connection_status = QtCore.pyqtSignal(bool)
    def __init__(self):
        super().__init__()
        self.ser = None
        self.running = True
    def find_serial_port(self):
        if not serial:
            return None
        ports = list(serial.tools.list_ports.comports())
        for p in ports:
            if 'XIAO' in p.description or 'USB Serial' in p.description:
                return p.device
        return None
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
        self.combo_boxes = []
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
    def get_processes(self):
        # Prioritize processes with active audio sessions at the top of the list, then add all other running processes
        import psutil
        names = ['System Master Volume']
        seen = set()
        audio_names = []
        if AudioUtilities:
            sessions = AudioUtilities.GetAllSessions()
            for session in sessions:
                if session.Process:
                    pname = session.Process.name()
                    if pname and pname.lower() not in seen:
                        audio_names.append(pname)
                        seen.add(pname.lower())
        # Add all running processes (not just audio sessions)
        for proc in psutil.process_iter(['name']):
            pname = proc.info['name']
            if pname and pname.lower() not in seen:
                names.append(pname)
                seen.add(pname.lower())
        # Prioritize audio sessions, then others
        return ['System Master Volume'] + sorted(audio_names, key=str.lower) + sorted([n for n in names if n not in audio_names and n != 'System Master Volume'], key=str.lower)
    def assign_process(self, fader_idx, combo_idx):
        procs = self.processes[self.page * 5:(self.page + 1) * 5]
        if combo_idx < len(procs):
            self.fader_assignments[fader_idx] = procs[combo_idx]
            self.fader_labels[fader_idx].setText(procs[combo_idx])
            self.save_settings()
    def set_app_volume(self, idx, val):
        proc = self.fader_assignments[idx]
        if proc == 'System Master Volume' and AudioUtilities:
            # Set master volume using IAudioEndpointVolume
            from pycaw.pycaw import IAudioEndpointVolume
            import comtypes
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, comtypes.CLSCTX_ALL, None)
            master = comtypes.cast(interface, comtypes.POINTER(IAudioEndpointVolume))
            master.SetMasterVolumeLevelScalar(val/127.0, None)
        elif proc and AudioUtilities and ISimpleAudioVolume:
            sessions = AudioUtilities.GetAllSessions()
            for session in sessions:
                if session.Process and proc.lower() in session.Process.name().lower():
                    volume_obj = session._ctl.QueryInterface(ISimpleAudioVolume)
                    volume_obj.SetMasterVolume(val/127.0, None)
        self.oled_previews[idx].setText(f'{proc[:8] if proc else "---"}: {int(val/127.0*100)}%')
        self.save_settings()
        self.update_led_bar()
    def update_faders(self, vals):
        for i, val in enumerate(vals):
            self.fader_sliders[i].setValue(val)
            proc = self.fader_assignments[i]
            self.oled_previews[i].setText(f'{proc[:8] if proc else "---"}: {int(val/127.0*100)}%')
        self.update_led_bar()
    def update_led_bar(self):
        # LED bar at bottom: show current main volume as a bar
        if hasattr(self, 'led_bar'):
            val = self.fader_sliders[0].value()
            percent = int(val / 127.0 * 100)
            self.led_bar.setValue(percent)
    def closeEvent(self, event):
        self.serial_worker.stop()
        self.thread.quit()
        self.thread.wait()
        event.accept()
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
        # Update process list to show all processes
        self.process_list.clear()
        self.process_list.addItems(self.get_processes())
    def save_settings(self):
        data = {
            'fader_assignments': self.fader_assignments,
            'fader_values': [slider.value() for slider in self.fader_sliders],
            'page': self.page
        }
        with open(SETTINGS_FILE, 'w') as f:
            json.dump(data, f)
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
            except Exception:
                pass
        else:
            # Default: only main fader assigned
            self.fader_assignments = ['System Master Volume'] + [None] * (len(self.fader_assignments) - 1)
    def update_status(self, connected):
        if connected:
            self.status_label.setText('Device: Connected')
            self.status_label.setStyleSheet('color: #08d9d6; font-weight: bold;')
        else:
            self.status_label.setText('Device: Not Connected')
            self.status_label.setStyleSheet('color: #ff2e63; font-weight: bold;')
    def paint_led_sim(self, widget, percent):
        # Simulate 8 LEDs as rectangles
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

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    # Timer to update LED simulation
    timer = QtCore.QTimer()
    timer.timeout.connect(lambda: window.led_sim.update())
    timer.start(50)
    window.show()
    sys.exit(app.exec_())
