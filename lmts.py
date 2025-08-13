#!/usr/bin/env python3
"""
LiveMicToSpeaker - Full runnable script with auto-stop-on-runtime (idle timeout removed).

Features:
 - GUI to choose input/output, settings in both main GUI and tray
 - Hotkey recorder (press keys to capture) and global registration via keyboard
 - Auto-start (HKCU Run) toggle
 - Auto-stop after user-specified runtime (when stream started)
 - Tray icon with toggle/settings/exit
 - Config file stored on system drive or AppData when frozen
"""

import sys
import os
import json
import time
import threading

import sounddevice as sd
import keyboard
import winreg

from PyQt5.QtWidgets import (
    QApplication, QWidget, QComboBox, QPushButton, QVBoxLayout, QLabel,
    QMessageBox, QLineEdit, QCheckBox, QHBoxLayout, QDialog, QSizePolicy
)
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QTimer, Qt

from pystray import Icon, Menu, MenuItem
from PIL import Image

# -------------------------
# Config path handling
# -------------------------
if getattr(sys, 'frozen', False):
    CONFIG_DIR = os.path.join(os.getenv('LOCALAPPDATA'), 'LiveMicToSpeaker')
else:
    SYSTEM_DRIVE = os.getenv('SYSTEMDRIVE', 'C:') + os.sep
    CONFIG_DIR = os.path.join(SYSTEM_DRIVE, 'LiveMicToSpeaker')

os.makedirs(CONFIG_DIR, exist_ok=True)
CONFIG_FILE = os.path.join(CONFIG_DIR, 'audio_config.json')

# -------------------------
# Globals and locks
# -------------------------
stream = None
stream_lock = threading.Lock()
tray_icon = None
mute_state = False
hotkey_handle = None

# Timer for auto-stop (when user starts stream)
auto_stop_timer = None
auto_stop_lock = threading.Lock()

# colored icons (simple solid squares)
ICON_IDLE = Image.new('RGB', (64, 64), color=(255, 0, 0))   # red = idle
ICON_ACTIVE = Image.new('RGB', (64, 64), color=(0, 200, 0)) # green = active

# -------------------------
# Device listing (PortAudio filter)
# -------------------------
def list_filtered_devices():
    """
    Returns (input_devices, output_devices) as lists of (index, name).
    Filters out obvious virtual devices and deduplicates by name.
    """
    devices = sd.query_devices()
    input_devices = []
    output_devices = []
    seen_inputs = set()
    seen_outputs = set()

    for idx, dev in enumerate(devices):
        name = dev.get('name', '').strip()
        if not name:
            continue
        if any(skip in name for skip in (
            "Microsoft Sound Mapper", "Primary Sound", "Loopback", "VoiceMeeter", "VB-Audio"
        )):
            continue
        if dev.get('max_input_channels', 0) > 0 and name not in seen_inputs:
            input_devices.append((idx, name))
            seen_inputs.add(name)
        if dev.get('max_output_channels', 0) > 0 and name not in seen_outputs:
            output_devices.append((idx, name))
            seen_outputs.add(name)

    return input_devices, output_devices

# -------------------------
# Config helpers
# -------------------------
def save_config(data):
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2)
    except Exception as e:
        print("Failed to save config:", e)

def load_config():
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        # defaults
        return {
            'input_device': None,
            'output_device': None,
            'blocksize': 256,
            'samplerate': 44100,
            'autostart': False,
            'hotkey': 'ctrl+m',
            'auto_stop_minutes': 0  # 0 = disabled
        }

# -------------------------
# Autostart (HKCU Run)
# -------------------------
RUN_KEY_PATH = r"Software\Microsoft\Windows\CurrentVersion\Run"
RUN_VALUE_NAME = "LiveMicToSpeaker"

def get_exe_path():
    return sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(sys.argv[0])

def add_to_startup():
    try:
        path = get_exe_path()
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY_PATH, 0, winreg.KEY_SET_VALUE)
        winreg.SetValueEx(key, RUN_VALUE_NAME, 0, winreg.REG_SZ, f'"{path}"')
        winreg.CloseKey(key)
        return True
    except Exception as e:
        print("add_to_startup error:", e)
        return False

def remove_from_startup():
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY_PATH, 0, winreg.KEY_SET_VALUE)
        try:
            winreg.DeleteValue(key, RUN_VALUE_NAME)
        except FileNotFoundError:
            pass
        winreg.CloseKey(key)
        return True
    except Exception as e:
        print("remove_from_startup error:", e)
        return False

# -------------------------
# Auto-stop timer helpers
# -------------------------
def _start_auto_stop_timer(minutes):
    """Starts/Restarts a single-shot timer that stops the stream after `minutes` minutes.
       If minutes <= 0 then no timer is started."""
    global auto_stop_timer
    with auto_stop_lock:
        # cancel previous
        if auto_stop_timer is not None:
            try:
                auto_stop_timer.cancel()
            except Exception:
                pass
            auto_stop_timer = None

        if minutes is None:
            return
        try:
            m = float(minutes)
        except Exception:
            return
        if m <= 0:
            return
        # create timer
        def _auto_stop_action():
            print(f"[auto_stop] runtime {m} minutes reached -> stop_stream()")
            stop_stream()
        auto_stop_timer = threading.Timer(m * 60.0, _auto_stop_action)
        auto_stop_timer.daemon = True
        auto_stop_timer.start()
        print(f"[auto_stop] timer started for {m} minutes")

def _cancel_auto_stop_timer():
    global auto_stop_timer
    with auto_stop_lock:
        if auto_stop_timer is not None:
            try:
                auto_stop_timer.cancel()
            except Exception:
                pass
            auto_stop_timer = None
            print("[auto_stop] timer canceled")

# -------------------------
# Audio stream control
# -------------------------
def start_stream():
    global stream, tray_icon
    with stream_lock:
        if stream or mute_state:
            return
        cfg = load_config()
        if cfg.get('input_device') is None or cfg.get('output_device') is None:
            print("No input/output configured - cannot start stream.")
            return

        def callback(indata, outdata, frames, t, status):
            if status:
                print("Stream status:", status)
            try:
                in_ch = indata.shape[1]
                out_ch = outdata.shape[1]
            except Exception:
                outdata[:] = indata
                return
            if in_ch == out_ch:
                outdata[:] = indata
            elif in_ch < out_ch:
                outdata[:, :in_ch] = indata
                outdata[:, in_ch:] = 0
            else:
                outdata[:] = indata[:, :out_ch]

        try:
            stream = sd.Stream(
                device=(cfg['input_device'], cfg['output_device']),
                callback=callback,
                samplerate=cfg.get('samplerate', 44100),
                blocksize=cfg.get('blocksize', 256),
                latency='low'
            )
            stream.start()
            if tray_icon:
                tray_icon.icon = ICON_ACTIVE
            print("[stream] started")
            # start auto-stop timer (if configured)
            _start_auto_stop_timer(cfg.get('auto_stop_minutes', 0))
        except Exception as e:
            print("Error starting stream:", e)
            stream = None

def stop_stream():
    global stream, tray_icon
    with stream_lock:
        if stream:
            try:
                stream.stop()
                stream.close()
            except Exception as e:
                print("Error stopping stream:", e)
            stream = None
            if tray_icon:
                tray_icon.icon = ICON_IDLE
            print("[stream] stopped")
        # cancel any auto-stop timer
        _cancel_auto_stop_timer()

def toggle_mute():
    global mute_state
    mute_state = not mute_state
    if mute_state:
        stop_stream()
    else:
        start_stream()

# -------------------------
# Hotkey management
# -------------------------
def register_hotkey(hotkey_str):
    """
    Register global hotkey using keyboard module. Returns True on success.
    Keeps track of hotkey_handle and removes previous registration.
    """
    global hotkey_handle
    try:
        if hotkey_handle is not None:
            try:
                keyboard.remove_hotkey(hotkey_handle)
            except Exception:
                pass
    except Exception:
        pass
    hotkey_handle = None
    if not hotkey_str:
        return False
    try:
        handle = keyboard.add_hotkey(hotkey_str, toggle_mute)
        hotkey_handle = handle
        print("Hotkey registered:", hotkey_str)
        return True
    except Exception as e:
        print("Failed to register hotkey:", e)
        return False

def run_hotkey():
    cfg = load_config()
    register_hotkey(cfg.get('hotkey', 'ctrl+m'))

# -------------------------
# Hotkey capture dialog (grabs keyboard on focus)
# -------------------------
class HotkeyCaptureDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Record Hotkey")
        self.setModal(True)
        self.setFixedSize(420, 130)
        self.captured = None
        self._grabbed = False

        self.label = QLabel("Now press the key combination you want to use.\n(Press Esc to cancel)", self)
        self.label.setWordWrap(True)
        self.preview = QLabel("", self)
        self.preview.setStyleSheet("font-weight: bold; font-size: 14px; padding-top:8px;")
        self.btn_cancel = QPushButton("Cancel")
        self.btn_cancel.clicked.connect(self.reject)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_layout.addWidget(self.btn_cancel)

        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.preview)
        layout.addLayout(btn_layout)
        self.setLayout(layout)

    def showEvent(self, ev):
        super().showEvent(ev)
        try:
            self.raise_()
            self.activateWindow()
            self.setFocus(Qt.ActiveWindowFocusReason)
            self.grabKeyboard()
            self._grabbed = True
        except Exception as e:
            print("grabKeyboard failed:", e)

    def closeEvent(self, ev):
        if self._grabbed:
            try:
                self.releaseKeyboard()
            except Exception:
                pass
        super().closeEvent(ev)

    def keyPressEvent(self, event):
        mods = []
        m = event.modifiers()
        if m & Qt.ControlModifier: mods.append("ctrl")
        if m & Qt.ShiftModifier:   mods.append("shift")
        if m & Qt.AltModifier:     mods.append("alt")
        if m & Qt.MetaModifier:    mods.append("win")

        key = event.key()

        if key == Qt.Key_Escape:
            self.captured = None
            if self._grabbed:
                try: self.releaseKeyboard()
                except: pass
            self.reject()
            return

        key_name = None
        if Qt.Key_A <= key <= Qt.Key_Z:
            key_name = chr(key).lower()
        elif Qt.Key_0 <= key <= Qt.Key_9:
            key_name = chr(key)
        elif Qt.Key_F1 <= key <= Qt.Key_F35:
            key_name = f"f{key - Qt.Key_F1 + 1}"
        else:
            special_map = {
                Qt.Key_Space: "space", Qt.Key_Tab: "tab", Qt.Key_Backspace: "backspace",
                Qt.Key_Return: "enter", Qt.Key_Enter: "enter", Qt.Key_Left: "left",
                Qt.Key_Right: "right", Qt.Key_Up: "up", Qt.Key_Down: "down",
                Qt.Key_Insert: "insert", Qt.Key_Delete: "delete", Qt.Key_Home: "home",
                Qt.Key_End: "end", Qt.Key_PageUp: "pageup", Qt.Key_PageDown: "pagedown",
                Qt.Key_Plus: "plus", Qt.Key_Minus: "minus", Qt.Key_Comma: "comma",
                Qt.Key_Period: "dot", Qt.Key_Slash: "slash", Qt.Key_Backslash: "backslash",
                Qt.Key_Semicolon: "semicolon", Qt.Key_QuoteLeft: "grave",
            }
            key_name = special_map.get(key, None)

        if key_name is None:
            txt = event.text()
            if txt:
                key_name = txt.lower()

        if not key_name:
            self.preview.setText("Unrecognized key, try again...")
            return

        parts = []
        for mname in ("ctrl", "alt", "shift", "win"):
            if mname in mods:
                parts.append(mname)
        parts.append(key_name)
        hotkey_str = "+".join(parts)
        self.preview.setText(hotkey_str)
        self.captured = hotkey_str
        if self._grabbed:
            try: self.releaseKeyboard()
            except: pass
        self.accept()

    def get_hotkey(self):
        return self.captured

# -------------------------
# Settings window
# -------------------------
class SettingsWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("LMTS Settings")
        self.setMinimumSize(360, 300)
        self.resize(360, 300)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        cfg = load_config()

        layout.addWidget(QLabel("Hotkey (press Record to capture):"))
        self.hotkey_input = QLineEdit(cfg.get('hotkey', 'ctrl+m'))
        btn_record = QPushButton("Record Hotkey")
        btn_record.clicked.connect(self.record_hotkey)

        hotrow = QHBoxLayout()
        hotrow.addWidget(self.hotkey_input, 3)
        hotrow.addWidget(btn_record, 1)
        layout.addLayout(hotrow)

        layout.addWidget(QLabel("Blocksize (e.g., 128,256):"))
        self.blocksize = QLineEdit(str(cfg.get('blocksize', 256)))
        self.blocksize.setMinimumHeight(28)
        self.blocksize.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        layout.addWidget(self.blocksize)

        layout.addWidget(QLabel("Sample Rate (e.g., 44100):"))
        self.samplerate = QLineEdit(str(cfg.get('samplerate', 44100)))
        self.samplerate.setMinimumHeight(28)
        self.samplerate.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        layout.addWidget(self.samplerate)

        layout.addWidget(QLabel("Auto-stop (minutes, 0 = disabled):"))
        self.auto_stop = QLineEdit(str(cfg.get('auto_stop_minutes', 0)))
        self.auto_stop.setMinimumHeight(28)
        self.auto_stop.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        layout.addWidget(self.auto_stop)

        self.autostart = QCheckBox("Auto-start with Windows")
        self.autostart.setChecked(cfg.get('autostart', False))
        layout.addWidget(self.autostart)

        save_btn = QPushButton("Save")
        save_btn.setMinimumHeight(34)
        save_btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        save_btn.clicked.connect(self.save)

        save_layout = QHBoxLayout()
        save_layout.addStretch()
        save_layout.addWidget(save_btn, 2)
        save_layout.addStretch()
        layout.addLayout(save_layout, stretch=1)

        note = QLabel("Note: avoid system-critical combos (e.g. Ctrl+Alt+Del).")
        note.setWordWrap(True)
        layout.addWidget(note)

        self.setLayout(layout)

    def record_hotkey(self):
        dlg = HotkeyCaptureDialog(self)
        if dlg.exec_() == QDialog.Accepted:
            hk = dlg.get_hotkey()
            if hk:
                self.hotkey_input.setText(hk)
            else:
                QMessageBox.warning(self, "No hotkey", "No hotkey captured.")

    def save(self):
        cfg = load_config()
        try:
            cfg['blocksize'] = int(self.blocksize.text())
            cfg['samplerate'] = int(self.samplerate.text())
            cfg['auto_stop_minutes'] = float(self.auto_stop.text())
        except ValueError:
            QMessageBox.critical(self, "Error", "Blocksize, Sample Rate and Auto-stop must be numbers.")
            return

        new_hotkey = self.hotkey_input.text().strip()
        if not new_hotkey:
            QMessageBox.critical(self, "Error", "Hotkey cannot be empty.")
            return

        cfg['hotkey'] = new_hotkey
        cfg['autostart'] = self.autostart.isChecked()
        save_config(cfg)

        # Apply autostart
        if cfg['autostart']:
            add_to_startup()
        else:
            remove_from_startup()

        # Re-register hotkey immediately
        if register_hotkey(new_hotkey):
            QMessageBox.information(self, "Saved", f"Settings saved. Hotkey: {new_hotkey}")
        else:
            QMessageBox.warning(self, "Saved (hotkey failed)",
                                "Settings saved but could not register hotkey. Try running as Administrator or choose a different combo.")

        # If stream running, restart auto-stop timer with new config
        if stream:
            _start_auto_stop_timer(cfg.get('auto_stop_minutes', 0))

    def closeEvent(self, event):
        # hide instead of close so app stays in tray
        self.hide()
        event.ignore()

# -------------------------
# Main GUI
# -------------------------
class MainApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Live Mic to Speaker")
        try:
            # If you bundle an icon, set QIcon(resource_path('app_icon.ico'))
            self.setWindowIcon(QIcon(sys.executable))
        except Exception:
            pass

        self.setMinimumSize(300, 160)
        self.resize(300, 160)

        self.input_devices, self.output_devices = list_filtered_devices()
        self.settings_window = None
        self.init_ui()
        self.load_cfg()

    def init_ui(self):
        layout = QVBoxLayout()

        layout.addWidget(QLabel("Mic Input:"))
        self.cb_in = QComboBox()
        self.cb_in.setMinimumHeight(28)
        self.cb_in.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        for idx, name in self.input_devices:
            self.cb_in.addItem(name, idx)
        layout.addWidget(self.cb_in)

        layout.addWidget(QLabel("Speaker Output:"))
        self.cb_out = QComboBox()
        self.cb_out.setMinimumHeight(28)
        self.cb_out.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        for idx, name in self.output_devices:
            self.cb_out.addItem(name, idx)
        layout.addWidget(self.cb_out)

        btn_layout = QHBoxLayout()
        self.btn_save = QPushButton("Start")
        btn_settings = QPushButton("Settings")
        self.btn_save.clicked.connect(self.save_start)
        btn_settings.clicked.connect(self.open_settings)

        self.btn_save.setMinimumHeight(34)
        btn_settings.setMinimumHeight(34)
        self.btn_save.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        btn_settings.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        btn_layout.addWidget(self.btn_save)
        btn_layout.addWidget(btn_settings)
        layout.addLayout(btn_layout, stretch=1)

        layout.addStretch()
        self.setLayout(layout)

    def load_cfg(self):
        cfg = load_config()
        for i, (idx, _) in enumerate(self.input_devices):
            if idx == cfg.get('input_device'):
                self.cb_in.setCurrentIndex(i)
        for i, (idx, _) in enumerate(self.output_devices):
            if idx == cfg.get('output_device'):
                self.cb_out.setCurrentIndex(i)

    def save_start(self):
        cfg = load_config()
        cfg['input_device'] = self.cb_in.currentData()
        cfg['output_device'] = self.cb_out.currentData()
        save_config(cfg)
        self.hide()
        start_stream()
        # change button text to Stop while running (user can re-open GUI to see state)
        self.btn_save.setText("Stop")

    def open_settings(self):
        if not self.settings_window:
            self.settings_window = SettingsWindow()
        self.settings_window.show()
        self.settings_window.activateWindow()

    def closeEvent(self, event):
        self.hide()
        event.ignore()

# -------------------------
# Tray icon
# -------------------------
def create_tray_icon(app_window):
    def on_toggle(icon, item):
        if stream:
            stop_stream()
            # reflect text if main window is alive
            try:
                app_window.btn_save.setText("Start")
            except Exception:
                pass
        else:
            start_stream()
            try:
                app_window.btn_save.setText("Stop")
            except Exception:
                pass
        icon.icon = ICON_ACTIVE if stream and not mute_state else ICON_IDLE

    def on_settings(icon, item):
        QTimer.singleShot(0, app_window.open_settings)

    def on_exit(icon, item):
        try:
            global hotkey_handle
            if hotkey_handle is not None:
                keyboard.remove_hotkey(hotkey_handle)
        except Exception:
            pass
        stop_stream()
        icon.stop()
        os._exit(0)

    menu = Menu(
        MenuItem('Toggle Mic', on_toggle, default=True),
        MenuItem('Settings', on_settings),
        MenuItem('Exit', on_exit)
    )
    icon = Icon("LiveMicToSpeaker", ICON_IDLE, "LiveMicToSpeaker", menu)
    threading.Thread(target=icon.run, daemon=True).start()
    return icon

# -------------------------
# Entry point
# -------------------------
def main():
    os.environ["QT_QPA_PLATFORM"] = "windows"
    app = QApplication(sys.argv)
    try:
        app.setWindowIcon(QIcon(sys.executable))
    except Exception:
        pass

    cfg = load_config()
    main_win = MainApp()
    main_win.show()

    # If devices set, hide main quickly (tray-only start)
    if cfg.get('input_device') is not None and cfg.get('output_device') is not None:
        QTimer.singleShot(400, main_win.hide)

    if cfg.get('autostart', False):
        add_to_startup()

    global tray_icon
    tray_icon = create_tray_icon(main_win)

    # Register initial hotkey (from config)
    run_hotkey()

    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
