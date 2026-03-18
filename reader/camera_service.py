"""
camera_service.py
------------------------
Minimal camera worker with a background thread that feeds frames to Dynamsoft.
Opens cameras strictly by *name* on Windows/macOS (no Linux).
"""
from __future__ import annotations

import os
import re
import sys
import time
import shutil
import threading
import subprocess
from pathlib import Path
from dataclasses import dataclass
from typing import Optional, Callable, List, Tuple, Dict

import cv2
from PIL import Image
from dynamsoft_barcode_reader_bundle import (
    CaptureVisionRouter, ImageSourceAdapter, ImageData, FileImageTag,
    EnumImagePixelFormat, EnumPresetTemplate, EnumBufferOverflowProtectionMode,
    EnumColourChannelUsageType, MultiFrameResultCrossFilter,
    EnumCapturedResultItemType, CapturedResultReceiver,
    DecodedBarcodesResult, ImageTag, Quadrilateral,
    EnumErrorCode, LicenseManager
)

try:
    import zxingcpp
except Exception:
    zxingcpp = None

cv2.setNumThreads(1)


# ---------------------- Dynamsoft glue ----------------------------------------
class MyCapturedResultReceiver(CapturedResultReceiver):
    def __init__(
        self,
        on_decoded: Optional[Callable[[str, str], None]] = None,
        on_detected: Optional[Callable[[Quadrilateral, str], None]] = None,
        on_error: Optional[Callable[[str], None]] = None,
        qr_validator: Optional[Callable[[str], bool]] = None
    ):
        super().__init__()
        self._on_decoded  = on_decoded
        self._on_detected = on_detected
        self._on_error    = on_error
        self._qr_validator = qr_validator

    def on_decoded_barcodes_received(self, result: DecodedBarcodesResult) -> None:
        if result.get_error_code() == EnumErrorCode.EC_UNSUPPORTED_JSON_KEY_WARNING:
            print("Warning:", result.get_error_string())
        elif result.get_error_code() != EnumErrorCode.EC_OK:
            print("Error:", result.get_error_string())

        items = result.get_items()
        if len(items) == 0:
            return

        tag: ImageTag = result.get_original_image_tag()
        if tag is not None:
            print("ImageID:", tag.get_image_id())
        print("Decoded", len(items), "barcodes.")

        for i, item in enumerate(items):
            text = item.get_text()
            text_norm = (text or "").strip()
            fmt = item.get_format_string()
            fmt_norm = (fmt or "").replace("_", "").upper()
            print("Result", i + 1)
            print("Barcode Format:", fmt)
            print("Barcode Text:", text)
            print()
            handled = False

            if self._on_detected:
                try:
                    self._on_detected(item.get_location(), fmt_norm)
                except Exception as e:
                    print("on_detected callback error:", e)

            if fmt_norm == "DATAMATRIX":
                if self._on_decoded:
                    try:
                        self._on_decoded(text_norm, fmt_norm)
                        handled = True
                    except Exception as e:
                        print("on_decoded callback error:", e)
            elif fmt_norm == "QRCODE":
                if self._qr_validator:
                    try:
                        if not self._qr_validator(text_norm):
                            continue

                    except Exception as e:
                        print("qr_validator error:", e)
                        continue
                if self._on_decoded:
                    try:
                        self._on_decoded(text_norm, fmt_norm)
                        handled = True
                    except Exception as e:
                        print("on_decoded callback error:", e)
            
            if not handled and self._on_error:
                try:
                    self._on_error(text_norm)
                except Exception as e:
                    print("on_error callback error:", e)


class MyVideoFetcher(ImageSourceAdapter):
    def __init__(self):
        super().__init__()

    def has_next_image_to_fetch(self) -> bool:
        return True


# ---------------------- Settings ----------------------------------------------
@dataclass
class CameraSettings:
    width: int = 1280
    height: int = 720
    fps: int = 10


# ---------------------- Camera Service ----------------------------------------
class CameraService:
    def __init__(
        self,
        device_index: Optional[int] = None,  # deprecated (ignored)
        settings: Optional[CameraSettings] = None,
        on_decoded: Optional[Callable[[str, str], None]] = None,
        on_detected: Optional[Callable[[Quadrilateral, str], None]] = None,
        on_error: Optional[Callable[[str], None]] = None,
        on_fatal: Optional[Callable[[str], None]] = None,
        allowed_cameras: Optional[List[str]] = None,
        qr_validator: Optional[Callable[[str], bool]] = None
    ):
        self.settings = settings or CameraSettings()

        self._running = False
        self._thread: Optional[threading.Thread] = None
        self._lock = threading.RLock()
        self._vc: Optional[cv2.VideoCapture] = None
        self._reopen: bool = False
        self.idx: Optional[int] = None

        self._on_decoded = on_decoded
        self._on_detected = on_detected
        self._on_error   = on_error
        self._on_fatal   = on_fatal
        self._allowed_cameras = [s for s in (allowed_cameras or []) if isinstance(s, str) and s.strip()]
        self._qr_validator = qr_validator

        self._black_frames_counter = 0
        self._not_ok_frames_counter = 0
        self._fallback_frame_counter = 0
        self._fallback_frame_interval = 3
        self._fallback_last_qr_text = ""
        self._fallback_last_qr_ts = 0.0
        self._fallback_dedup_window = 5.0
        self._fallback_last_error_log = 0.0
        self._last_decoded_text = ""
        self._last_decoded_fmt = ""
        self._last_decoded_ts = 0.0
        self._decoded_dedup_window = 0.75

        # Dynamsoft license
        code, msg = LicenseManager.init_license(
            "f0068dAAAAOB8i2g/xXbP4obF58F6t2RDYovSbkg5GklRZO/By1UFLQ8yrjZgEQJ2zQkzRbbdMGfR+X4O9Y94e6J8lFq2qe0="
        )
        if code not in (EnumErrorCode.EC_OK.value, EnumErrorCode.EC_LICENSE_CACHE_USED.value):
            print("License initialization failed:", code, msg)
            raise RuntimeError(f"License initialization failed: {msg}")

    # ------------- lifecycle ---------------------------------------------------
    def start(self) -> None:
        with self._lock:
            if self._running:
                return
            self._running = True
            self._thread = threading.Thread(target=self._run, name="CameraService", daemon=True)
            self._thread.start()

    def stop(self) -> None:
        with self._lock:
            self._running = False
            if self._vc is not None:
                try: self._vc.release()
                except Exception: pass
                self._vc = None
        t = self._thread
        if t and t.is_alive():
            t.join(timeout=2.0)

    def close(self) -> None:
        if self._vc is not None:
            try: self._vc.release()
            except Exception: pass
            self._vc = None

    def reconnect (self) -> None:
        """Kept for compatibility with main.py; ignores parameters and forces a reopen."""
        with self._lock:
            self._reopen = True

    def get_index (self):
        return self.idx

    def _emit_decoded(self, text: str, fmt: str) -> None:
        if not self._on_decoded:
            return
        text_norm = (text or "").strip()
        fmt_norm = (fmt or "").replace("_", "").upper()
        if not text_norm:
            return

        now = time.monotonic()
        if (
            text_norm == self._last_decoded_text
            and fmt_norm == self._last_decoded_fmt
            and (now - self._last_decoded_ts) < self._decoded_dedup_window
        ):
            return

        self._last_decoded_text = text_norm
        self._last_decoded_fmt = fmt_norm
        self._last_decoded_ts = now
        self._on_decoded(text_norm, fmt_norm)

    @staticmethod
    def _extract_format_name(format_value: object) -> str:
        if format_value is None:
            return ""
        if hasattr(format_value, "name"):
            return str(getattr(format_value, "name") or "")
        return str(format_value)

    def _try_decode_qr_fallback(self, frame_bgr) -> None:
        """
        Optional QR fallback decoder using zxing-cpp.
        Runs every few frames and only emits validated QR payloads.
        """
        if zxingcpp is None or self._on_decoded is None:
            return

        self._fallback_frame_counter += 1
        if (self._fallback_frame_counter % self._fallback_frame_interval) != 0:
            return

        try:
            frame_rgb = cv2.cvtColor(frame_bgr, cv2.COLOR_BGR2RGB)
            pil_img = Image.fromarray(frame_rgb)
            for result in zxingcpp.read_barcodes(pil_img):
                text = (getattr(result, "text", "") or "").strip()
                if not text:
                    continue

                fmt_name = self._extract_format_name(getattr(result, "format", None))
                fmt_norm = fmt_name.replace("_", "").replace(" ", "").replace("-", "").upper()
                if "QRCODE" not in fmt_norm:
                    continue

                if self._qr_validator:
                    try:
                        if not self._qr_validator(text):
                            continue
                    except Exception as e:
                        print("qr_validator error (zxing fallback):", e)
                        continue

                now = time.monotonic()
                if (
                    text == self._fallback_last_qr_text
                    and (now - self._fallback_last_qr_ts) < self._fallback_dedup_window
                ):
                    continue

                self._fallback_last_qr_text = text
                self._fallback_last_qr_ts = now

                try:
                    self._emit_decoded(text, "QRCODE")
                except Exception as e:
                    print("on_decoded callback error (zxing fallback):", e)
                return
        except Exception as e:
            now = time.monotonic()
            if now - self._fallback_last_error_log >= 5.0:
                print("zxing fallback decode error:", e)
                self._fallback_last_error_log = now

    # ---------------- helpers (platform) --------------------------------------
    @staticmethod
    def _norm(s: str) -> str:
        return re.sub(r"\s+", " ", s or "").strip().casefold()

    @staticmethod
    def _ensure_ffmpeg_on_path() -> None:
        """If ffmpeg is not on PATH, try to add ./bin or script dir."""
        if shutil.which("ffmpeg"):
            return
        here = os.path.abspath(os.path.dirname(__file__))
        cwd  = os.path.abspath(os.getcwd())
        candidates = [
            os.path.join(here, "bin"),
            os.path.join(cwd, "bin"),
            here,
            cwd,
            os.environ.get("FFMPEG_DIR", "") or "",
            os.environ.get("FFMPEG_PATH", "") or "",
        ]
        for d in candidates:
            if not d:
                continue
            for exe in ("ffmpeg.exe", "ffmpeg"):
                if os.path.exists(os.path.join(d, exe)):
                    os.environ["PATH"] = d + os.pathsep + os.environ.get("PATH", "")
                    return

    @staticmethod
    def _config_cap(cap: cv2.VideoCapture, settings: CameraSettings) -> None:
        try: cap.set(cv2.CAP_PROP_BUFFERSIZE, 1)
        except Exception: pass
        if sys.platform.startswith("win"):
            # Many webcams behave better in MJPG on Windows
            try: cap.set(cv2.CAP_PROP_FOURCC, cv2.VideoWriter_fourcc(*"MJPG"))
            except Exception: pass
        if getattr(settings, "width", None):  cap.set(cv2.CAP_PROP_FRAME_WIDTH,  settings.width)
        if getattr(settings, "height", None): cap.set(cv2.CAP_PROP_FRAME_HEIGHT, settings.height)
        if getattr(settings, "fps", None):    cap.set(cv2.CAP_PROP_FPS,          settings.fps)

    # ---------- Windows (DirectShow) name → index via ffmpeg -------------------
    @classmethod
    def _dshow_index_map(cls) -> Tuple[List[str], Dict[str, int]]:
        """
        Returns:
          ordered: friendly names (video) in DShow order → index 0..N-1
          name2idx: dict (casefold) mapping friendly & alternative names to index
        """
        cls._ensure_ffmpeg_on_path()
        ordered: List[str] = []
        name2idx: Dict[str, int] = {}
        if not shutil.which("ffmpeg"):
            return ordered, name2idx
        try:
            _exe = "ffmpeg"
            _base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
            _bin = _base / "bin" / ("ffmpeg.exe" if sys.platform.startswith("win") else "ffmpeg")

            # Evitar finestra de consola a Windows
            _extra = {}
            if sys.platform.startswith("win"):
                _si = subprocess.STARTUPINFO()
                _si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                _extra["startupinfo"] = _si
                _extra["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)

            p = subprocess.run(
                [_exe, "-hide_banner", "-f", "dshow", "-list_devices", "true", "-i", "dummy"],
                capture_output=True, text=True, timeout=3.0, **_extra
            )

            txt = (p.stderr or "") + (p.stdout or "")
            in_videos = False
            last_idx: Optional[int] = None
            for raw in txt.splitlines():
                line = raw.strip()
                if "DirectShow video devices" in line:
                    in_videos = True
                    continue
                if in_videos and "DirectShow audio devices" in line:
                    break

                m_name = re.search(r'"([^"]+)"\s*\(video\)', line)
                if m_name:
                    fname = m_name.group(1).strip()
                    last_idx = len(ordered)
                    ordered.append(fname)
                    name2idx[fname.casefold()] = last_idx
                    continue

                if last_idx is not None:
                    m_alt = re.search(r'Alternative name\s+"([^"]+)"', line)
                    if m_alt:
                        alt = m_alt.group(1).strip()
                        name2idx[alt.casefold()] = last_idx
        except Exception:
            pass
        return ordered, name2idx

    # ---------- macOS (AVFoundation) name → index via ffmpeg -------------------
    @classmethod
    def _mac_find_index_for_name(cls, name: str) -> Optional[int]:
        cls._ensure_ffmpeg_on_path()
        if not shutil.which("ffmpeg"):
            return None
        try:
            _exe = "ffmpeg"
            _base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
            _bin = _base / "bin" / ("ffmpeg.exe" if sys.platform.startswith("win") else "ffmpeg")
            if _bin.exists():
                _exe = str(_bin)

            # No calen flags a macOS; passar _extra buit és inofensiu
            _extra = {}

            proc = subprocess.run(
                [_exe, "-hide_banner", "-f", "avfoundation", "-list_devices", "true", "-i", ""],
                capture_output=True, text=True, timeout=10.0, **_extra
            )

            txt = (proc.stderr or "") + (proc.stdout or "")
            candidates: List[Tuple[int, str]] = []
            for line in txt.splitlines():
                m = re.search(r"\[(\d+)\]\s+(.+)$", line.strip())
                if not m:
                    continue
                idx, nm = int(m.group(1)), m.group(2).strip()
                candidates.append((idx, nm))
            # exact first, then partial
            for idx, nm in candidates:
                if cls._norm(nm) == cls._norm(name):
                    return idx
            for idx, nm in candidates:
                if cls._norm(name) in cls._norm(nm):
                    return idx
        except Exception:
            pass
        return None

    # ---------- Open helpers ---------------------------------------------------
    def _open_capture_by_index(self, index: int, retries: int = 3, retry_delay: float = 0.25) -> Optional[cv2.VideoCapture]:
        """Internal: open by numeric index (picked from mapping)."""
        if sys.platform.startswith("win"):
            backends = [cv2.CAP_DSHOW, cv2.CAP_MSMF, cv2.CAP_ANY]
        elif sys.platform == "darwin":
            backends = [cv2.CAP_AVFOUNDATION, cv2.CAP_ANY]
        else:
            backends = [cv2.CAP_V4L2, cv2.CAP_ANY]

        for _ in range(max(1, retries)):
            for be in backends:
                cap = None
                try:
                    cap = cv2.VideoCapture(index, be)
                except Exception:
                    cap = None

                if not cap or not cap.isOpened():
                    if cap is not None:
                        try: cap.release()
                        except Exception: pass
                    continue

                self._config_cap(cap, self.settings)

                ok, _ = cap.read()
                if ok:
                    self.idx = index
                    print(f"[CameraService] Opened numeric index {index} (backend={be})")
                    return cap

                try: cap.release()
                except Exception: pass
            time.sleep(retry_delay)
        return None

    def _open_capture_by_name_strict(self, retries: int = 2, retry_delay: float = 0.25) -> Optional[cv2.VideoCapture]:
        """Open ONLY cameras whose name is in self._allowed_cameras (first match wins)."""
        names = [n for n in self._allowed_cameras if isinstance(n, str) and n.strip()]
        if not names:
            return None

        if sys.platform.startswith("win"):
            ordered, name2idx = self._dshow_index_map()
            print(f"[CameraService] DShow order: {ordered}")

            # 1) exact match on any allowed name → index
            for nm in names:
                self.idx = name2idx.get(nm.casefold())
                if self.idx is not None:
                    return self._open_capture_by_index(self.idx, retries=retries, retry_delay=retry_delay)
            # 2) partial match (first in DShow order)
            for i, fname in enumerate(ordered):
                for nm in names:
                    if self._norm(nm) in self._norm(fname):
                        self.idx = i
                        return self._open_capture_by_index(i, retries=retries, retry_delay=retry_delay)
            return None

        elif sys.platform == "darwin":
            # Try each allowed name until one maps to an AVFoundation index
            for nm in names:
                self.idx = self._mac_find_index_for_name(nm)
                if self.idx is not None:
                    return self._open_capture_by_index(self.idx, retries=retries, retry_delay=retry_delay)
            return None

        else:
            return None

    def _open_capture_any(self, retries: int = 2, retry_delay: float = 0.25, max_index: int = 8) -> Optional[cv2.VideoCapture]:
        """Best-effort fallback: open the first available camera index."""
        for idx in range(max(1, max_index)):
            cap = self._open_capture_by_index(idx, retries=retries, retry_delay=retry_delay)
            if cap is not None:
                return cap
        return None

    def _open_capture_preferred(self, retries: int = 2, retry_delay: float = 0.25) -> Optional[cv2.VideoCapture]:
        """
        Try strict eBuho camera names first; if unavailable, fallback to any camera.
        This allows local testing without the eBuho hardware connected.
        """
        if self._allowed_cameras:
            cap = self._open_capture_by_name_strict(retries=retries, retry_delay=retry_delay)
            if cap is not None:
                return cap
            print(f"[CameraService] Preferred cameras not found ({self._allowed_cameras}). Falling back to first available camera.")

        return self._open_capture_any(retries=retries, retry_delay=retry_delay)

    # ----------------------- main thread --------------------------------------
    def _run(self) -> None:
        cvr = CaptureVisionRouter()

        fetcher = MyVideoFetcher()
        fetcher.set_max_image_count(2)
        fetcher.set_buffer_overflow_protection_mode(EnumBufferOverflowProtectionMode.BOPM_UPDATE)
        fetcher.set_colour_channel_usage_type(EnumColourChannelUsageType.CCUT_AUTO)
        cvr.set_input(fetcher)

        mf = MultiFrameResultCrossFilter()
        mf.enable_result_cross_verification(EnumCapturedResultItemType.CRIT_BARCODE, True)
        mf.enable_result_deduplication(EnumCapturedResultItemType.CRIT_BARCODE, True)
        mf.set_duplicate_forget_time(EnumCapturedResultItemType.CRIT_BARCODE, 5000)
        cvr.add_result_filter(mf)

        receiver = MyCapturedResultReceiver(self._emit_decoded, self._on_detected, self._on_error, self._qr_validator)
        cvr.add_result_receiver(receiver)

        code, msg = cvr.start_capturing(EnumPresetTemplate.PT_READ_BARCODES, False)
        if code != EnumErrorCode.EC_OK:
            print("Error starting capture:", msg)
            return

        image_id = 0
        target_fps = int(self.settings.fps) if getattr(self.settings, "fps", None) else 10
        target_interval = max(0.04, 1.0 / max(1, target_fps))
        self._last_tick = time.monotonic()

        try:
            # Initial open — preferred list first, fallback to any camera.
            self._vc = self._open_capture_preferred()

            if self._vc is None:
                print(f"[CameraService] Unable to open camera (preferred={self._allowed_cameras})")
                try: cvr.stop_capturing(False, True)
                except Exception: pass
                if self._on_fatal:
                    self._on_fatal("Unable to open camera.")
                return

            frame_ctr = 0
            while True:
                with self._lock:
                    if not self._running:
                        break
                    do_switch = self._reopen
                    if do_switch:
                        self._reopen = False
                        old_cap = self._vc
                        self._vc = None
                    else:
                        old_cap = None

                if do_switch:
                    if old_cap is not None:
                        try:
                            old_cap.release()
                            time.sleep(1.0)
                        except Exception:
                            pass

                    time.sleep(0.15)  # let driver settle

                    new_cap = self._open_capture_preferred(retries=3, retry_delay=0.25)
                    if not new_cap:
                        self._reopen = True

                    with self._lock:
                        self._vc = new_cap

                if self._vc is None:
                    time.sleep(0.05)
                    continue

                ok, frame = self._vc.read()
                if not ok or frame is None:
                    self._not_ok_frames_counter += 1
                    if self._not_ok_frames_counter >= 3:
                        self._reopen = True
                        self._not_ok_frames_counter = 0
                    time.sleep(0.01)
                    continue
                else:
                    self._not_ok_frames_counter = 0

                    proc_frame = frame
                    frame_ctr += 1

                    # Trigger: 5 frames totalment negres -> reopen
                    if frame.mean() == 0:
                        self._black_frames_counter += 1
                        if self._black_frames_counter >= 5:
                            self._reopen = True
                            self._black_frames_counter = 0
                        else:
                            time.sleep(0.5)

                    else:
                        self._black_frames_counter = 0
                        self._try_decode_qr_fallback(frame)

                        if (frame_ctr % 2) != 0:
                            time.sleep(0.005)
                            proc_frame = cv2.bitwise_not(frame)

                        image_id += 1

                        tag = FileImageTag("", 0, 0)
                        tag.set_image_id(image_id)

                        h, w = proc_frame.shape[0], proc_frame.shape[1]
                        step = proc_frame.strides[0]

                        image = ImageData(
                            proc_frame.tobytes(), w, h, step,
                            EnumImagePixelFormat.IPF_BGR_888, 0, tag
                        )
                        fetcher.add_image_to_buffer(image)

                        # Throttle to target FPS
                        now = time.monotonic()
                        elapsed = now - self._last_tick
                        if elapsed < target_interval:
                            time.sleep(target_interval - elapsed)
                        self._last_tick = time.monotonic()

                        time.sleep(0.002)

        finally:
            try:
                cvr.stop_capturing(False, True)
            except Exception:
                pass
            if self._vc is not None:
                try: self._vc.release()
                except Exception: pass
                self._vc = None
