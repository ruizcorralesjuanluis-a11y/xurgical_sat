"""
serial_service.py
-----------------
SerialService: auto-connect to configured port, hot-switch at runtime, and simple send_text API.
Uses pyserial; non-blocking writes; optional background reader (disabled by default).
"""
from __future__ import annotations
import threading, time
from dataclasses import dataclass
from typing import Optional, Callable

try:
    import serial  # pyserial
    import serial.tools.list_ports as list_ports
except Exception:
    serial = None
    list_ports = None


@dataclass
class SerialSettings:
    baudrate: int = 115200
    timeout: float = 0.1      # seconds
    write_timeout: float = 0.25


class SerialService:
    def __init__(self, port: Optional[str] = None, settings: Optional[SerialSettings] = None):
        self._lock = threading.RLock()
        self._ser: Optional["serial.Serial"] = None if serial else None
        self._running = False
        self._thread: Optional[threading.Thread] = None
        self.port = port                       # e.g. "COM5" or "/dev/tty.usbserial-XXXX"
        self.settings = settings or SerialSettings()
        self._want_port = port
        self._auto_connect = True
        self._on_line: Optional[Callable[[str], None]] = None  # optional reader callback

    # --- lifecycle ---
    def start(self, start_reader: bool = False) -> None:
        with self._lock:
            if self._running:
                return
            self._running = True
            if start_reader:
                self._thread = threading.Thread(target=self._reader_loop, name="SerialService", daemon=True)
                self._thread.start()
        # try auto-connect once on start
        if self._auto_connect and self.port and serial:
            self._open_port(self.port)

    def stop(self) -> None:
        with self._lock:
            self._running = False
        t = self._thread
        if t and t.is_alive():
            t.join(timeout=1.0)
        self._close_port()

    def close(self) -> None:
        self.stop()

    # --- configuration / hot switch ---
    def switch_port(self, new_port: Optional[str]) -> None:
        """Hot-switch to a new port. If None, just close."""
        with self._lock:
            self._want_port = new_port
        # perform switch outside lock
        if new_port:
            self._open_port(new_port)
        else:
            self._close_port()

    # --- public API ---
    def send_text(self, text: str, append_newline: bool = True) -> bool:
        """Send a line of text over serial. Returns True if written."""
        if text is None:
            return False
        data = (text + ("\n" if append_newline else "")).encode("utf-8", errors="replace")
        with self._lock:
            if not self._ser or not self._ser.is_open:
                # try to open desired port if set
                if self._auto_connect and self._want_port:
                    self._open_port(self._want_port)
                if not self._ser or not self._ser.is_open:
                    return False
            try:
                self._ser.write_timeout = self.settings.write_timeout
                self._ser.write(data)
                self._ser.flush()
                return True
            except Exception:
                # on error, drop connection; caller may retry later
                self._close_port()
                return False

    def set_on_line(self, cb: Optional[Callable[[str], None]]) -> None:
        with self._lock:
            self._on_line = cb

    # --- internals ---
    def _open_port(self, port: str) -> None:
        if not serial:
            return
        with self._lock:
            # if already on this port and open, nothing to do
            if self._ser and self._ser.is_open and self.port == port:
                return
            # close old first
            if self._ser:
                try:
                    self._ser.close()
                except Exception:
                    pass
                self._ser = None
            try:
                self._ser = serial.Serial(
                    port=port,
                    baudrate=self.settings.baudrate,
                    timeout=self.settings.timeout,
                    write_timeout=self.settings.write_timeout,
                )
                self.port = port
            except Exception:
                # keep closed on failure
                self._ser = None

    def _close_port(self) -> None:
        with self._lock:
            if self._ser:
                try:
                    self._ser.close()
                except Exception:
                    pass
                self._ser = None

    def _reader_loop(self) -> None:
        # Optional background reader (disabled by default)
        while True:
            with self._lock:
                if not self._running:
                    break
                ser = self._ser
                cb = self._on_line
            if ser and ser.is_open:
                try:
                    line = ser.readline().decode("utf-8", errors="replace").rstrip("\r\n")
                    if line and cb:
                        cb(line)
                except Exception:
                    self._close_port()
                    time.sleep(0.1)
            else:
                time.sleep(0.05)


# Convenience utility
def list_serial_ports_ebuho() -> list[str]:
    if not list_ports:
        return []
    # Usamos VID/PID específicos del eBuho Reader mencionados en su main.py
    return [p.device for p in list_ports.comports() if getattr(p, "vid", None) == 4617 and getattr(p, "pid", None) == 60161]
