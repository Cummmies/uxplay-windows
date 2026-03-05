import os
import sys
import logging
import shlex
import subprocess
import threading
import time
import winreg
import webbrowser
import ctypes
import tkinter as tk
from tkinter import scrolledtext
from pathlib import Path
from typing import List, Optional

import pystray
from PIL import Image

# ─── Constants ────────────────────────────────────────────────────────────────

APP_NAME    = "uxplay-windows"
APPDATA_DIR = Path(os.environ["APPDATA"]) / "uxplay-windows"
LOG_FILE    = APPDATA_DIR / f"{APP_NAME}.log"

APPDATA_DIR.mkdir(parents=True, exist_ok=True)

# ─── Log Window ───────────────────────────────────────────────────────────────

class LogWindow:
    """
    Dark-themed floating log viewer. Opens from the tray menu.
    Buffers all messages so history is shown even if opened late.
    Runs its own Tk mainloop in a dedicated thread so it never
    blocks the tray or the server.
    """

    BG         = "#1e1e1e"
    FG         = "#d4d4d4"
    FG_INFO    = "#9cdcfe"
    FG_WARN    = "#dcdcaa"
    FG_ERROR   = "#f44747"
    FG_DEBUG   = "#808080"
    FONT       = ("Consolas", 9)

    def __init__(self):
        self._root: Optional[tk.Tk] = None
        self._text: Optional[scrolledtext.ScrolledText] = None
        self._autoscroll: Optional[tk.BooleanVar] = None
        self._buffer: List[tuple] = []          # (formatted_msg, levelname)
        self._lock   = threading.Lock()
        self._thread: Optional[threading.Thread] = None

    # ── public ────────────────────────────────────────────────────────────────

    def append(self, message: str, level: str = "INFO") -> None:
        """Thread-safe: buffer the line and push to the text widget if open."""
        with self._lock:
            self._buffer.append((message, level))
        # schedule onto the Tk event loop if window is alive
        if self._root is not None:
            try:
                self._root.after(0, self._insert, message, level)
            except Exception:
                pass

    def show(self) -> None:
        """Called from the tray menu (non-main thread). Opens or raises the window."""
        if self._thread and self._thread.is_alive():
            # already running — just raise it
            if self._root:
                try:
                    self._root.after(0, self._raise)
                except Exception:
                    pass
            return
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def destroy(self) -> None:
        if self._root:
            try:
                self._root.after(0, self._root.destroy)
            except Exception:
                pass

    # ── internal ──────────────────────────────────────────────────────────────

    def _raise(self):
        try:
            self._root.deiconify()
            self._root.lift()
            self._root.focus_force()
        except Exception:
            pass

    def _run(self):
        """Runs entirely on the log-window thread."""
        self._root = tk.Tk()
        self._root.title(f"{APP_NAME}  —  Live Log")
        self._root.configure(bg=self.BG)
        self._root.geometry("860x420")
        self._root.resizable(True, True)
        self._root.protocol("WM_DELETE_WINDOW", self._on_close)

        # ── status bar at top ─────────────────────────────────────────────────
        self._status_var = tk.StringVar(value="Starting…")
        status_bar = tk.Label(
            self._root, textvariable=self._status_var,
            bg="#007acc", fg="#ffffff",
            font=("Segoe UI", 9, "bold"),
            anchor="w", padx=8, pady=3
        )
        status_bar.pack(fill=tk.X, side=tk.TOP)

        # ── toolbar ───────────────────────────────────────────────────────────
        toolbar = tk.Frame(self._root, bg="#2d2d2d", pady=3)
        toolbar.pack(fill=tk.X, side=tk.TOP)

        btn = dict(bg="#3c3c3c", fg=self.FG, relief=tk.FLAT,
                   font=("Segoe UI", 9), padx=10, pady=2,
                   activebackground="#505050", activeforeground="#fff",
                   cursor="hand2", bd=0)

        tk.Button(toolbar, text="🗑  Clear",     command=self._clear,         **btn).pack(side=tk.LEFT, padx=(6, 2))
        tk.Button(toolbar, text="📂  Open Log File", command=self._open_file, **btn).pack(side=tk.LEFT, padx=2)
        tk.Button(toolbar, text="📋  Copy All",  command=self._copy_all,      **btn).pack(side=tk.LEFT, padx=2)

        self._autoscroll = tk.BooleanVar(value=True)
        tk.Checkbutton(
            toolbar, text="Auto-scroll",
            variable=self._autoscroll,
            bg="#2d2d2d", fg=self.FG,
            selectcolor="#3c3c3c",
            activebackground="#2d2d2d",
            activeforeground=self.FG,
            font=("Segoe UI", 9)
        ).pack(side=tk.RIGHT, padx=8)

        # ── text area ─────────────────────────────────────────────────────────
        frame = tk.Frame(self._root, bg=self.BG)
        frame.pack(fill=tk.BOTH, expand=True, padx=4, pady=(2, 0))

        hbar = tk.Scrollbar(frame, orient=tk.HORIZONTAL, bg="#2d2d2d")
        hbar.pack(side=tk.BOTTOM, fill=tk.X)

        self._text = scrolledtext.ScrolledText(
            frame,
            bg=self.BG, fg=self.FG,
            font=self.FONT,
            wrap=tk.NONE,
            state=tk.DISABLED,
            insertbackground=self.FG,
            selectbackground="#264f78",
            relief=tk.FLAT,
            borderwidth=0,
            xscrollcommand=hbar.set,
        )
        self._text.pack(fill=tk.BOTH, expand=True)
        hbar.config(command=self._text.xview)

        # colour tags
        self._text.tag_configure("INFO",    foreground=self.FG_INFO)
        self._text.tag_configure("WARNING", foreground=self.FG_WARN)
        self._text.tag_configure("ERROR",   foreground=self.FG_ERROR)
        self._text.tag_configure("DEBUG",   foreground=self.FG_DEBUG)
        self._text.tag_configure("DEFAULT", foreground=self.FG)

        # replay buffered lines
        with self._lock:
            snapshot = list(self._buffer)
        for msg, lvl in snapshot:
            self._insert(msg, lvl)

        self._status_var.set(f"Log file: {LOG_FILE}")
        self._root.mainloop()
        self._root = None
        self._text = None
        self._autoscroll = None

    def _insert(self, message: str, level: str) -> None:
        if not self._text:
            return
        tag = level if level in ("INFO", "WARNING", "ERROR", "DEBUG") else "DEFAULT"
        self._text.configure(state=tk.NORMAL)
        self._text.insert(tk.END, message + "\n", tag)
        self._text.configure(state=tk.DISABLED)
        if self._autoscroll and self._autoscroll.get():
            self._text.see(tk.END)

    def _clear(self) -> None:
        with self._lock:
            self._buffer.clear()
        if self._text:
            self._text.configure(state=tk.NORMAL)
            self._text.delete("1.0", tk.END)
            self._text.configure(state=tk.DISABLED)

    def _open_file(self) -> None:
        try:
            os.startfile(str(LOG_FILE))
        except Exception:
            pass

    def _copy_all(self) -> None:
        if not self._text or not self._root:
            return
        content = self._text.get("1.0", tk.END)
        self._root.clipboard_clear()
        self._root.clipboard_append(content)

    def _on_close(self) -> None:
        # hide rather than destroy so it reopens instantly next time
        if self._root:
            self._root.withdraw()

    def update_status(self, text: str) -> None:
        """Update the blue status bar at the top (e.g. 'UxPlay running — PID 1234')."""
        if self._root and hasattr(self, "_status_var"):
            try:
                self._root.after(0, lambda: self._status_var.set(text))
            except Exception:
                pass


# ── Logging handler that feeds lines into LogWindow ───────────────────────────

class _GUIHandler(logging.Handler):
    def __init__(self, window: LogWindow):
        super().__init__()
        self._window = window

    def emit(self, record: logging.LogRecord) -> None:
        try:
            self._window.append(self.format(record), record.levelname)
        except Exception:
            pass


# ─── Logging Setup ────────────────────────────────────────────────────────────

_log_window = LogWindow()
_fmt        = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")

_fh = logging.FileHandler(str(LOG_FILE), encoding="utf-8")
_fh.setFormatter(_fmt)

_sh = logging.StreamHandler(sys.stdout)
_sh.setFormatter(_fmt)

_gh = _GUIHandler(_log_window)
_gh.setFormatter(_fmt)

logging.basicConfig(level=logging.INFO, handlers=[_fh, _sh, _gh])

# ─── Silent Audio Session ─────────────────────────────────────────────────────
# winmm plays a silent loop inside this process → one audio session in the
# volume mixer. MediaPlayer (WinRT) is muted and only used for the SMTC handle.

_SILENT_WAV = (
    b'RIFF$\x00\x00\x00WAVEfmt \x10\x00\x00\x00\x01\x00\x01\x00'
    b'\x44\xac\x00\x00\x88X\x01\x00\x02\x00\x10\x00data\x00\x00\x00\x00'
)
import tempfile as _tf
_wav_tmp = _tf.NamedTemporaryFile(suffix=".wav", delete=False)
_wav_tmp.write(_SILENT_WAV)
_wav_tmp.close()
_WAV_PATH = _wav_tmp.name

_winmm         = ctypes.windll.winmm
_SND_FILENAME  = 0x20000
_SND_ASYNC     = 0x0001
_SND_LOOP      = 0x0008
_SND_NODEFAULT = 0x0002

def _audio_start():
    _winmm.PlaySoundW(_WAV_PATH, None, _SND_FILENAME | _SND_ASYNC | _SND_LOOP | _SND_NODEFAULT)

def _audio_stop():
    _winmm.PlaySoundW(None, None, 0)

# ─── Windows Media Session (SMTC) ─────────────────────────────────────────────

class MediaSessionManager:
    on_play:  Optional[callable] = None
    on_pause: Optional[callable] = None
    on_stop:  Optional[callable] = None

    def __init__(self):
        self._controls  = None
        self._updater   = None
        self._available = False
        self._MPS       = None
        self._setup()

    def _setup(self):
        try:
            from winsdk.windows.media import MediaPlaybackStatus, MediaPlaybackType
            from winsdk.windows.media.playback import MediaPlayer

            self._player          = MediaPlayer()
            self._player.is_muted = True
            self._player.volume   = 0.0

            self._controls = self._player.system_media_transport_controls
            self._updater  = self._controls.display_updater

            self._controls.is_enabled       = True
            self._controls.is_play_enabled  = True
            self._controls.is_pause_enabled = True
            self._controls.is_stop_enabled  = True
            self._controls.add_button_pressed(self._on_button)

            self._updater.type = MediaPlaybackType.VIDEO
            self._updater.video_properties.title    = "AirPlay"
            self._updater.video_properties.subtitle = "uxplay-windows"
            self._updater.update()

            self._MPS       = MediaPlaybackStatus
            self._available = True
            logging.info("SMTC registered OK")

        except ImportError:
            logging.warning("winsdk not installed — SMTC overlay disabled (pip install winsdk)")
        except Exception:
            logging.exception("SMTC setup failed")

    def _update(self, status, title: str):
        if not self._available:
            return
        try:
            self._controls.playback_status = status
            self._updater.video_properties.title = title
            self._updater.update()
        except Exception:
            logging.exception("SMTC update failed")

    def set_playing(self):
        self._update(self._MPS.PLAYING, "AirPlay — Connected")
        _audio_start()
        logging.info("SMTC → PLAYING")

    def set_stopped(self):
        self._update(self._MPS.STOPPED, "AirPlay")
        _audio_stop()
        logging.info("SMTC → STOPPED")

    def set_paused(self):
        self._update(self._MPS.PAUSED, "AirPlay — Paused")
        _audio_stop()
        logging.info("SMTC → PAUSED")

    def _on_button(self, sender, args):
        # 0=Play 1=Pause 2=Stop 3=Prev 4=Next
        try:
            btn = args.button
            logging.info("SMTC button: %s", btn)
            if   btn == 0 and self.on_play:  self.on_play()
            elif btn == 1 and self.on_pause: self.on_pause()
            elif btn == 2 and self.on_stop:  self.on_stop()
        except Exception:
            logging.exception("SMTC button error")

# ─── Path Discovery ───────────────────────────────────────────────────────────

class Paths:
    def __init__(self):
        if getattr(sys, "frozen", False):
            cand = Path(sys._MEIPASS) if hasattr(sys, "_MEIPASS") else Path(sys.executable).parent
        else:
            cand = Path(__file__).resolve().parent

        internal         = cand / "_internal"
        self.resource_dir = internal if internal.is_dir() else cand
        self.icon_file    = self.resource_dir / "icon.ico"

        ux1 = self.resource_dir / "bin" / "uxplay.exe"
        ux2 = self.resource_dir / "uxplay.exe"
        self.uxplay_exe = ux1 if ux1.exists() else ux2

        self.appdata_dir    = APPDATA_DIR
        self.arguments_file = APPDATA_DIR / "arguments.txt"

# ─── Argument File Manager ────────────────────────────────────────────────────

class ArgumentManager:
    def __init__(self, file_path: Path):
        self.file_path = file_path

    def ensure_exists(self) -> None:
        self.file_path.parent.mkdir(parents=True, exist_ok=True)
        if not self.file_path.exists():
            self.file_path.write_text("", encoding="utf-8")
            logging.info("Created arguments.txt at %s", self.file_path)

    def read_args(self) -> List[str]:
        if not self.file_path.exists():
            return []
        text = self.file_path.read_text(encoding="utf-8").strip()
        if not text:
            return []
        try:
            return shlex.split(text)
        except ValueError as e:
            logging.error("Could not parse arguments.txt: %s", e)
            return []

# ─── Server Process Manager ──────────────────────────────────────────────────

class ServerManager:
    def __init__(self, exe_path: Path, arg_mgr: ArgumentManager,
                 media_mgr: Optional[MediaSessionManager] = None,
                 log_window: Optional[LogWindow] = None):
        self.exe_path   = exe_path
        self.arg_mgr    = arg_mgr
        self.media_mgr  = media_mgr
        self.log_window = log_window
        self.process: Optional[subprocess.Popen] = None
        self._on_state_change: Optional[callable] = None

    def is_running(self) -> bool:
        return self.process is not None and self.process.poll() is None

    def _set_status(self, text: str):
        if self.log_window:
            self.log_window.update_status(text)

    def start(self) -> None:
        if self.is_running():
            logging.info("UxPlay already running (PID %s)", self.process.pid)
            return
        if not self.exe_path.exists():
            logging.error("uxplay.exe not found at %s", self.exe_path)
            self._set_status(f"❌  uxplay.exe not found at {self.exe_path}")
            return

        cmd = [str(self.exe_path)] + self.arg_mgr.read_args()
        logging.info("Starting UxPlay: %s", cmd)
        try:
            self.process = subprocess.Popen(cmd, creationflags=subprocess.CREATE_NO_WINDOW)
            logging.info("UxPlay started (PID %s)", self.process.pid)
            self._set_status(f"✅  UxPlay running — PID {self.process.pid}")
            if self.media_mgr:
                self.media_mgr.set_playing()
            if self._on_state_change:
                self._on_state_change()
            # watch for unexpected exit in background
            threading.Thread(target=self._watch, daemon=True).start()
        except Exception:
            logging.exception("Failed to launch UxPlay")
            self._set_status("❌  Failed to launch UxPlay — see log")

    def stop(self) -> None:
        if not self.is_running():
            logging.info("UxPlay not running.")
            return
        pid = self.process.pid
        logging.info("Stopping UxPlay (PID %s)…", pid)
        try:
            self.process.terminate()
            self.process.wait(timeout=3)
            logging.info("UxPlay stopped.")
        except subprocess.TimeoutExpired:
            logging.warning("Timeout — killing PID %s", pid)
            self.process.kill()
            self.process.wait()
        except Exception:
            logging.exception("Error stopping UxPlay")
        finally:
            self.process = None
            self._set_status("⏹  UxPlay stopped")
            if self.media_mgr:
                self.media_mgr.set_stopped()
            if self._on_state_change:
                self._on_state_change()

    def _watch(self):
        """Detect if uxplay.exe exits on its own and update state."""
        if self.process:
            self.process.wait()
            if not self.is_running():   # wasn't stopped intentionally
                logging.warning("UxPlay exited unexpectedly (code %s)", self.process.returncode if self.process else "?")
                self._set_status("⚠️  UxPlay exited unexpectedly — check log")
                self.process = None
                if self.media_mgr:
                    self.media_mgr.set_stopped()
                if self._on_state_change:
                    self._on_state_change()

# ─── Auto-Start Manager ───────────────────────────────────────────────────────

class AutoStartManager:
    RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"

    def __init__(self, app_name: str, exe_cmd: str):
        self.app_name = app_name
        self.exe_cmd  = exe_cmd

    def is_enabled(self) -> bool:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.RUN_KEY, 0, winreg.KEY_READ) as k:
                val, _ = winreg.QueryValueEx(k, self.app_name)
                return self.exe_cmd in val
        except FileNotFoundError:
            return False
        except Exception:
            logging.exception("Autostart check failed")
            return False

    def enable(self):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.RUN_KEY, 0, winreg.KEY_SET_VALUE) as k:
                winreg.SetValueEx(k, self.app_name, 0, winreg.REG_SZ, self.exe_cmd)
            logging.info("Autostart enabled")
        except Exception:
            logging.exception("Failed to enable Autostart")

    def disable(self):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.RUN_KEY, 0, winreg.KEY_SET_VALUE) as k:
                winreg.DeleteValue(k, self.app_name)
            logging.info("Autostart disabled")
        except FileNotFoundError:
            pass
        except Exception:
            logging.exception("Failed to disable Autostart")

    def toggle(self):
        self.disable() if self.is_enabled() else self.enable()

# ─── System Tray Icon UI ─────────────────────────────────────────────────────

class TrayIcon:
    def __init__(self, icon_path, server_mgr, arg_mgr, auto_mgr, log_window):
        self.server_mgr = server_mgr
        self.arg_mgr    = arg_mgr
        self.auto_mgr   = auto_mgr
        self.log_window = log_window

        def start_label(_):
            return "⬤  Running" if server_mgr.is_running() else "Start UxPlay"

        menu = pystray.Menu(
            pystray.MenuItem(start_label,         lambda _: server_mgr.start(),
                             enabled=lambda _: not server_mgr.is_running()),
            pystray.MenuItem("Stop UxPlay",        lambda _: server_mgr.stop(),
                             enabled=lambda _: server_mgr.is_running()),
            pystray.MenuItem("Restart UxPlay",     lambda _: self._restart()),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Show Log",           lambda _: log_window.show()),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(
                "Autostart with Windows",
                lambda _: auto_mgr.toggle(),
                checked=lambda _: auto_mgr.is_enabled()
            ),
            pystray.MenuItem("Edit UxPlay Arguments", lambda _: self._open_args()),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("License",
                lambda _: webbrowser.open(
                    "https://github.com/leapbtw/uxplay-windows/blob/main/LICENSE.md")),
            pystray.MenuItem("Exit",               lambda _: self._exit()),
        )

        self.icon = pystray.Icon(
            name=APP_NAME,
            icon=Image.open(icon_path),
            title=APP_NAME,
            menu=menu,
        )

        server_mgr._on_state_change = self._refresh

    def _refresh(self):
        try:
            self.icon.update_menu()
        except Exception:
            pass

    def _restart(self):
        logging.info("Restarting UxPlay")
        self.server_mgr.stop()
        self.server_mgr.start()

    def _open_args(self):
        self.arg_mgr.ensure_exists()
        try:
            os.startfile(str(self.arg_mgr.file_path))
        except Exception:
            logging.exception("Could not open arguments.txt")

    def _exit(self):
        logging.info("Exiting uxplay-windows")
        self.server_mgr.stop()
        self.log_window.destroy()
        _audio_stop()
        try:
            os.unlink(_WAV_PATH)
        except Exception:
            pass
        self.icon.stop()

    def run(self):
        self.icon.run()

# ─── Application Orchestration ───────────────────────────────────────────────

class Application:
    def __init__(self):
        self.paths      = Paths()
        self.arg_mgr    = ArgumentManager(self.paths.arguments_file)
        self.media_mgr  = MediaSessionManager()

        exe_cmd = (f'"{sys.executable}"' if getattr(sys, "frozen", False)
                   else f'"{sys.executable}" "{Path(__file__).resolve()}"')

        self.auto_mgr   = AutoStartManager(APP_NAME, exe_cmd)
        self.server_mgr = ServerManager(
            self.paths.uxplay_exe, self.arg_mgr,
            self.media_mgr, _log_window
        )
        self.tray = TrayIcon(
            self.paths.icon_file,
            self.server_mgr, self.arg_mgr,
            self.auto_mgr, _log_window
        )

        self.media_mgr.on_play  = self.server_mgr.start
        self.media_mgr.on_pause = self.server_mgr.stop
        self.media_mgr.on_stop  = self.server_mgr.stop

    def run(self):
        self.arg_mgr.ensure_exists()
        logging.info("uxplay-windows starting up")
        logging.info("uxplay.exe path: %s", self.paths.uxplay_exe)
        logging.info("Log file: %s", LOG_FILE)
        threading.Thread(target=self._delayed_start, daemon=True).start()
        self.tray.run()
        logging.info("Tray exited — shutting down")

    def _delayed_start(self):
        time.sleep(3)
        self.server_mgr.start()

if __name__ == "__main__":
    Application().run()
