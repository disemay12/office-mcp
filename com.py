"""Shared COM plumbing for the Office MCP bridge.

All Office automation runs on one dedicated STA thread. Office applications are
apartment-bound COM servers; calling them from arbitrary thread-pool threads is
unreliable, and a COM object created in one thread cannot be used in another.
The dispatcher keeps one app instance per host alive, serializes every call and
guards against two MCP server processes (two chats) driving Office at once.
"""

from __future__ import annotations

import concurrent.futures
import datetime as dt
import decimal
import math
import os
import queue
import re
import threading
import time
from contextlib import suppress
from typing import Any, Callable

import pythoncom
import pywintypes
import win32com.client as win32
import win32event

APP_PROGIDS = {
    "word": "Word.Application",
    "excel": "Excel.Application",
    "powerpoint": "PowerPoint.Application",
}

DEFAULT_TIMEOUT = 180.0
MUTEX_NAME = r"Local\OfficeMCP_Bridge_Serial"
MUTEX_WAIT_MS = 60_000
MUTEX_WAIT_ABANDONED = 0x80

# Relative paths resolve here; outputs created without an explicit path land here.
BASE_DIR = os.path.abspath(
    os.environ.get("OFFICE_MCP_BASE_DIR")
    or os.path.join(os.path.expanduser("~"), "Documents", "Office MCP")
)

_BUSY_HRESULTS = {
    -2147418111,  # RPC_E_CALL_REJECTED
    -2147417846,  # RPC_E_SERVERCALL_RETRYLATER
}

# A cached application proxy is assumed alive. These HRESULTs mean the Office
# process went away, so the handle is dropped and the call is retried once.
_DISCONNECTED_HRESULTS = {
    -2147417848,  # RPC_E_DISCONNECTED
    -2147220995,  # CO_E_OBJNOTCONNECTED
    -2147023174,  # RPC_S_SERVER_UNAVAILABLE
    -2147221021,  # MK_E_UNAVAILABLE
    -2147418105,  # RPC_E_SERVER_DIED
    -2147417838,  # RPC_E_SERVER_DIED_DNE
}

# Per-process COM settings cache: progid -> (id(app), applied Visible value).
_APP_SETTINGS: dict[str, tuple[int, bool | None]] = {}


class OfficeError(RuntimeError):
    """User-facing failure raised by an Office tool."""


def _hresult(exc: BaseException) -> int | None:
    args = getattr(exc, "args", ())
    if args and isinstance(args[0], int):
        return args[0]
    return None


def _is_busy(exc: BaseException) -> bool:
    return _hresult(exc) in _BUSY_HRESULTS


def _is_disconnected(exc: BaseException) -> bool:
    return _hresult(exc) in _DISCONNECTED_HRESULTS


def _com_error_text(exc: BaseException) -> str:
    if isinstance(exc, pywintypes.com_error):
        hr = _hresult(exc)
        detail = exc.strerror or (exc.args[2] if len(exc.args) > 2 else None)
        parts = [str(detail or exc)]
        if hr is not None:
            parts.append(f"HRESULT 0x{hr & 0xFFFFFFFF:08X}")
        return " / ".join(parts)
    return str(exc)


def _retry_busy(fn: Callable[[], Any], attempts: int = 4, delay: float = 0.35) -> Any:
    for attempt in range(attempts):
        try:
            return fn()
        except pywintypes.com_error as exc:
            if not _is_busy(exc) or attempt == attempts - 1:
                raise
            time.sleep(delay * (attempt + 1))
    raise AssertionError("unreachable")


class _Dispatcher:
    def __init__(self) -> None:
        self._queue: queue.Queue[Any] = queue.Queue()
        self._apps: dict[str, Any] = {}
        self._dead = threading.Event()
        self._thread = threading.Thread(
            target=self._loop, name="office-com", daemon=True
        )
        self._thread.start()

    # -- worker ---------------------------------------------------------

    def _loop(self) -> None:
        try:
            pythoncom.CoInitialize()
        except BaseException as exc:  # pragma: no cover - startup failure
            self._fail_all(exc)
            return

        mutex = None
        with suppress(Exception):
            mutex = win32event.CreateMutex(None, False, MUTEX_NAME)

        try:
            while True:
                job = self._queue.get()
                if job is None:
                    return
                fn, future = job
                if not future.set_running_or_notify_cancel():
                    continue
                acquired = False
                try:
                    if mutex is not None:
                        rc = win32event.WaitForSingleObject(mutex, MUTEX_WAIT_MS)
                        if rc not in (0, MUTEX_WAIT_ABANDONED):
                            raise OfficeError(
                                "Another Office MCP session is busy with Office. "
                                "Wait for it to finish and retry."
                            )
                        acquired = True
                    future.set_result(fn())
                except BaseException as exc:  # noqa: BLE001 - forwarded to caller
                    with suppress(Exception):
                        future.set_exception(exc)
                finally:
                    if acquired:
                        with suppress(Exception):
                            win32event.ReleaseMutex(mutex)
        except BaseException as exc:  # pragma: no cover - fatal worker failure
            self._fail_all(exc)
        finally:
            self._apps.clear()
            with suppress(Exception):
                pythoncom.CoUninitialize()

    def _fail_all(self, exc: BaseException) -> None:
        self._dead.set()
        while True:
            try:
                job = self._queue.get_nowait()
            except queue.Empty:
                break
            if job is None:
                continue
            _, future = job
            with suppress(Exception):
                future.set_exception(
                    OfficeError(f"Office COM worker stopped: {exc}")
                )

    # -- public API -----------------------------------------------------

    def call(self, fn: Callable[[], Any], timeout: float = DEFAULT_TIMEOUT) -> Any:
        if self._dead.is_set() or not self._thread.is_alive():
            raise OfficeError("Office COM worker is not running. Restart the MCP server.")
        future: concurrent.futures.Future[Any] = concurrent.futures.Future()
        self._queue.put((fn, future))
        try:
            return future.result(timeout)
        except concurrent.futures.TimeoutError:
            future.cancel()
            raise OfficeError(
                f"Office did not answer within {timeout:.0f}s. A modal Office dialog is "
                "probably open (file in use, recovery, login). Close it and retry; the "
                "operation may still finish."
            ) from None

    def app(self, app_key: str, *, create: bool = True) -> Any:
        """Return a live application object. Runs on the COM thread."""
        progid = APP_PROGIDS.get(app_key)
        if progid is None:
            raise OfficeError(f"Unknown Office application: {app_key}")

        # Trust the cached proxy: probing it on every call costs a COM round
        # trip. run() drops the handle and retries once if it went stale.
        app = self._apps.get(progid)
        if app is not None:
            return app

        try:
            app = _retry_busy(lambda: win32.GetActiveObject(progid))
        except Exception:
            if not create:
                raise OfficeError(
                    f"{app_key} is not running. Start it or create a document first."
                ) from None
            try:
                app = win32.DispatchEx(progid)
            except pywintypes.com_error as exc:
                raise OfficeError(
                    f"Cannot start {app_key}: {_com_error_text(exc)}"
                ) from exc

        self._apps[progid] = app
        return app

    def forget(self, app_key: str) -> None:
        progid = APP_PROGIDS.get(app_key)
        if progid is not None:
            self._apps.pop(progid, None)
            _APP_SETTINGS.pop(progid, None)

    def shutdown(self) -> None:
        self._queue.put(None)
        self._thread.join(timeout=10)


DISPATCHER = _Dispatcher()


def _prepare(app_key: str, app: Any, visible: bool | None) -> None:
    """Apply non-intrusive settings once per application instance."""
    state = _APP_SETTINGS.get(app_key)
    if state is None or state[0] != id(app):
        if app_key == "word":
            with suppress(Exception):
                app.DisplayAlerts = 0
        elif app_key == "excel":
            with suppress(Exception):
                app.DisplayAlerts = False
            with suppress(Exception):
                app.AskToUpdateLinks = False
        elif app_key == "powerpoint":
            with suppress(Exception):
                app.DisplayAlerts = 1
        _APP_SETTINGS[app_key] = (id(app), None)
        applied_visible = None
    else:
        applied_visible = state[1]

    if visible is None or applied_visible == bool(visible):
        return
    try:
        app.Visible = bool(visible)
    except pywintypes.com_error:
        # PowerPoint refuses to hide its window; that is fine.
        if visible:
            raise
    else:
        _APP_SETTINGS[app_key] = (id(app), bool(visible))


def run(
    app_key: str,
    fn: Callable[[Any], Any],
    *,
    visible: bool | None = None,
    create: bool = True,
    timeout: float = DEFAULT_TIMEOUT,
) -> Any:
    """Execute ``fn(app)`` on the COM thread and return JSON-safe data."""

    def job() -> Any:
        app = DISPATCHER.app(app_key, create=create)
        _prepare(app_key, app, visible)
        return jsonable(fn(app))

    try:
        return DISPATCHER.call(job, timeout=timeout)
    except pywintypes.com_error as exc:
        if not _is_disconnected(exc):
            raise OfficeError(f"{app_key}: {_com_error_text(exc)}") from exc
        DISPATCHER.forget(app_key)
        try:
            return DISPATCHER.call(job, timeout=timeout)
        except pywintypes.com_error as retry_exc:
            raise OfficeError(f"{app_key}: {_com_error_text(retry_exc)}") from retry_exc


# -- paths --------------------------------------------------------------


def resolve_path(path: str, *, must_exist: bool = False) -> str:
    raw = str(path).strip().strip('"').strip("'")
    if not raw:
        raise OfficeError("Empty file path")
    expanded = os.path.expandvars(os.path.expanduser(raw))
    if not os.path.isabs(expanded):
        expanded = os.path.join(BASE_DIR, expanded)
    full = os.path.abspath(expanded)
    if must_exist and not os.path.exists(full):
        raise OfficeError(f"File not found: {full}")
    return full


def prepare_output(path: str) -> str:
    full = resolve_path(path)
    parent = os.path.dirname(full)
    if parent:
        os.makedirs(parent, exist_ok=True)
    return full


def default_output(app_key: str, ext: str) -> str:
    os.makedirs(BASE_DIR, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    return os.path.join(BASE_DIR, f"{app_key}-{stamp}.{ext.lstrip('.')}")


def file_info(path: str) -> dict[str, Any]:
    try:
        stat = os.stat(path)
    except OSError:
        return {"path": path, "exists": False}
    return {
        "path": path,
        "exists": True,
        "size_bytes": stat.st_size,
        "modified": dt.datetime.fromtimestamp(stat.st_mtime).isoformat(timespec="seconds"),
    }


# -- conversions --------------------------------------------------------


def cm_to_points(value: float) -> float:
    return float(value) * 28.3464567


def points_to_cm(value: float) -> float:
    return round(float(value) / 28.3464567, 2)


def hex_to_bgr(color: str | int | None) -> int | None:
    """Accept '#RRGGBB', 'RRGGBB' or an already-packed OLE colour integer."""
    if color is None:
        return None
    if isinstance(color, bool):
        raise OfficeError("Invalid colour value")
    if isinstance(color, int):
        return color
    text = str(color).strip().lstrip("#")
    if not re.fullmatch(r"[0-9a-fA-F]{6}", text):
        raise OfficeError(
            f"Invalid colour '{color}'. Use a hex string like '#1F4E79'."
        )
    r, g, b = int(text[0:2], 16), int(text[2:4], 16), int(text[4:6], 16)
    return r | (g << 8) | (b << 16)


def jsonable(value: Any) -> Any:
    if value is None or isinstance(value, (str, bool, int)):
        return value
    if isinstance(value, float):
        return value if math.isfinite(value) else None
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        return value.isoformat()
    if isinstance(value, dt.timedelta):
        return value.total_seconds()
    if isinstance(value, decimal.Decimal):
        return float(value)
    if isinstance(value, dict):
        return {str(k): jsonable(v) for k, v in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [jsonable(v) for v in value]
    if isinstance(value, bytes):
        return value.decode("utf-8", "replace")
    if isinstance(value, pywintypes.TimeType):
        return value.isoformat()
    try:
        return str(value)
    except Exception:  # pragma: no cover - defensive
        return repr(value)


def normalize_matrix(values: Any) -> list[list[Any]]:
    """Turn a scalar / row / matrix into a rectangular 2-D list."""
    if isinstance(values, dict):
        raise OfficeError("values must be a list, not an object")
    if isinstance(values, (str, bytes)) or not hasattr(values, "__iter__"):
        return [[values]]

    rows = list(values)
    if not rows:
        raise OfficeError("values is empty")
    if not any(isinstance(row, (list, tuple)) for row in rows):
        return [rows]

    matrix: list[list[Any]] = []
    width = 0
    for row in rows:
        cells = list(row) if isinstance(row, (list, tuple)) else [row]
        matrix.append(cells)
        width = max(width, len(cells))
    for row in matrix:
        row.extend([None] * (width - len(row)))
    return matrix


def is_index(value: Any) -> bool:
    return isinstance(value, str) and value.strip().isdigit()
