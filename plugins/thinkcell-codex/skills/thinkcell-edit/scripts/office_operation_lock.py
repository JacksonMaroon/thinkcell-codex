"""Fail-fast, cross-process serialization for plugin-owned Office operations.

This module deliberately knows nothing about PowerPoint, think-cell, COM, or
the UI.  Callers hold ``OfficeOperationLock`` around their complete native
operation.  A timed-out child is quarantined so a later invocation cannot
mistake an unreturned helper for an idle Office session.
"""
from __future__ import annotations

import ctypes
from ctypes import wintypes
import argparse
from functools import wraps
import inspect
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import threading
import time
from contextlib import AbstractContextManager
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any, Mapping, Sequence


WAIT_OBJECT_0 = 0
WAIT_ABANDONED = 0x80
WAIT_TIMEOUT = 0x102
INFINITE = 0xFFFFFFFF
ERROR_ALREADY_EXISTS = 183
ERROR_INVALID_PARAMETER = 87
STILL_ACTIVE = 259
MUTEX_NAME = "Local\\user-suppliedThinkcell.OfficeOperation.v1"


class OfficeOperationLockError(RuntimeError):
    """Base error with a JSON-safe diagnostic payload."""

    def __init__(self, message: str, diagnostic: Mapping[str, Any]):
        super().__init__(message)
        self.diagnostic = dict(diagnostic)


class OfficeOperationBusy(OfficeOperationLockError):
    """Another plugin process currently owns the native-operation mutex."""


class OfficeOperationQuarantined(OfficeOperationLockError):
    """A previously timed-out helper may still be operating Office."""


class OfficeOperationTimedOut(OfficeOperationLockError):
    """The launched child outlived its allowed time and was quarantined."""


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _state_path() -> Path:
    override = os.environ.get("user-supplied_THINKCELL_OFFICE_LOCK_STATE")
    if override:
        return Path(override).expanduser().resolve()
    base = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "user-suppliedThinkcell"
    return base / "office-operation-lock.json"


def _read_state(path: Path) -> dict[str, Any]:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except FileNotFoundError:
        return {}
    except (OSError, json.JSONDecodeError):
        # A diagnostic file must never turn into permission to operate Office.
        return {"corrupt_state": True, "path": str(path)}


def _write_state(path: Path, state: Mapping[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary = tempfile.mkstemp(prefix="office-lock-", suffix=".json", dir=path.parent)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            json.dump(state, handle, indent=2, sort_keys=True)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, path)
    finally:
        try:
            os.unlink(temporary)
        except FileNotFoundError:
            pass


def _creation_time(pid: int) -> int | None:
    """Return Windows FILETIME creation ticks, or None if the PID is gone."""
    kernel32 = _kernel32()
    process = kernel32.OpenProcess(0x1000, False, pid)  # PROCESS_QUERY_LIMITED_INFORMATION
    if not process:
        return None
    try:
        created = wintypes.FILETIME()
        exited = wintypes.FILETIME()
        kernel = wintypes.FILETIME()
        user = wintypes.FILETIME()
        if not kernel32.GetProcessTimes(process, ctypes.byref(created), ctypes.byref(exited),
                                        ctypes.byref(kernel), ctypes.byref(user)):
            return None
        return (created.dwHighDateTime << 32) | created.dwLowDateTime
    finally:
        kernel32.CloseHandle(process)


def _process_status(identity: Mapping[str, Any]) -> tuple[str, str]:
    """Return whether an exact process instance is alive, dead, or unknown.

    ``unknown`` deliberately fails closed: an access denial or an incomplete
    identity cannot prove that a potentially Office-touching helper exited.
    """
    pid = identity.get("pid")
    created = identity.get("creation_time")
    if not isinstance(pid, int) or not isinstance(created, int):
        return "unknown", "missing_pid_or_creation_time"
    kernel32 = _kernel32()
    process = kernel32.OpenProcess(0x1000, False, pid)
    if not process:
        error = ctypes.get_last_error()
        if error == ERROR_INVALID_PARAMETER:
            return "dead", "pid_not_found"
        return "unknown", "open_process_failed_%d" % error
    try:
        actual_filetime = wintypes.FILETIME()
        exited = wintypes.FILETIME()
        kernel = wintypes.FILETIME()
        user = wintypes.FILETIME()
        if not kernel32.GetProcessTimes(process, ctypes.byref(actual_filetime), ctypes.byref(exited),
                                        ctypes.byref(kernel), ctypes.byref(user)):
            return "unknown", "get_process_times_failed_%d" % ctypes.get_last_error()
        actual = (actual_filetime.dwHighDateTime << 32) | actual_filetime.dwLowDateTime
        if actual != created:
            return "dead", "pid_reused"
        exit_code = wintypes.DWORD()
        # A terminated child may retain a process handle in its parent.  It has
        # the same creation time but is no longer capable of touching Office.
        if not kernel32.GetExitCodeProcess(process, ctypes.byref(exit_code)):
            return "unknown", "get_exit_code_failed_%d" % ctypes.get_last_error()
        return ("alive", "still_active") if exit_code.value == STILL_ACTIVE else ("dead", "exited")
    finally:
        kernel32.CloseHandle(process)


def _same_process(identity: Mapping[str, Any]) -> bool:
    """Compatibility helper; recovery must use ``_process_status`` instead."""
    return _process_status(identity)[0] == "alive"


def _identity(operation: str, metadata: Mapping[str, Any] | None) -> dict[str, Any]:
    value: dict[str, Any] = {
        "pid": os.getpid(),
        "creation_time": _creation_time(os.getpid()),
        "thread_id": threading.get_ident(),
        "operation": operation,
        "started_at": _utc_now(),
    }
    if metadata:
        value["metadata"] = dict(metadata)
    return value


def _kernel32():
    """Bind pointer-sized Win32 handles correctly on 64-bit Python."""
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.CreateMutexW.argtypes = (wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR)
    kernel32.CreateMutexW.restype = wintypes.HANDLE
    kernel32.WaitForSingleObject.argtypes = (wintypes.HANDLE, wintypes.DWORD)
    kernel32.WaitForSingleObject.restype = wintypes.DWORD
    kernel32.ReleaseMutex.argtypes = (wintypes.HANDLE,)
    kernel32.ReleaseMutex.restype = wintypes.BOOL
    kernel32.CloseHandle.argtypes = (wintypes.HANDLE,)
    kernel32.CloseHandle.restype = wintypes.BOOL
    kernel32.OpenProcess.argtypes = (wintypes.DWORD, wintypes.BOOL, wintypes.DWORD)
    kernel32.OpenProcess.restype = wintypes.HANDLE
    kernel32.GetProcessTimes.argtypes = (wintypes.HANDLE, ctypes.POINTER(wintypes.FILETIME),
                                         ctypes.POINTER(wintypes.FILETIME), ctypes.POINTER(wintypes.FILETIME),
                                         ctypes.POINTER(wintypes.FILETIME))
    kernel32.GetProcessTimes.restype = wintypes.BOOL
    kernel32.GetExitCodeProcess.argtypes = (wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD))
    kernel32.GetExitCodeProcess.restype = wintypes.BOOL
    return kernel32


@dataclass
class _Held:
    handle: int
    thread_id: int
    depth: int
    owner: dict[str, Any]


_local_guard = threading.RLock()
_held: _Held | None = None


def serialized_office(function):
    """Serialize an Office-touching entrypoint while leaving explicit previews pure.

    Entrypoints either expose an ``execute`` argument or receive an argparse
    namespace as their first argument.  Functions without that switch are
    treated as live operations, which keeps older feature APIs protected.
    """
    signature = inspect.signature(function)

    @wraps(function)
    def wrapped(*args, **kwargs):
        bound = signature.bind(*args, **kwargs)
        bound.apply_defaults()
        execute = bound.arguments.get("execute")
        if not isinstance(execute, bool) and bound.arguments:
            execute = getattr(next(iter(bound.arguments.values())), "execute", None)
        if execute is False:
            return function(*args, **kwargs)
        operation = "%s.%s" % (function.__module__, function.__qualname__)
        with OfficeOperationLock(operation):
            return function(*args, **kwargs)

    return wrapped


class OfficeOperationLock(AbstractContextManager["OfficeOperationLock"]):
    """Own the plugin-wide Office mutex, failing immediately if it is busy."""

    def __init__(self, operation: str, *, metadata: Mapping[str, Any] | None = None):
        if not operation or not isinstance(operation, str):
            raise ValueError("operation must be a non-empty string")
        self.operation = operation
        self.metadata = metadata
        self._outermost = False

    @property
    def state_path(self) -> Path:
        return _state_path()

    def __enter__(self) -> "OfficeOperationLock":
        global _held
        current_thread = threading.get_ident()
        with _local_guard:
            if _held is not None and _held.thread_id == current_thread:
                _held.depth += 1
                return self
            kernel32 = _kernel32()
            handle = kernel32.CreateMutexW(None, False, MUTEX_NAME)
            if not handle:
                raise OSError(ctypes.get_last_error(), "CreateMutexW failed")
            result = kernel32.WaitForSingleObject(handle, 0)
            if result == WAIT_TIMEOUT:
                state = _read_state(self.state_path)
                kernel32.CloseHandle(handle)
                raise OfficeOperationBusy("Another plugin native operation is active", {
                    "status": "OFFICE_OPERATION_BUSY", "mutex": MUTEX_NAME,
                    "state_path": str(self.state_path), "owner": state.get("owner"),
                })
            if result not in (WAIT_OBJECT_0, WAIT_ABANDONED):
                error = ctypes.get_last_error()
                kernel32.CloseHandle(handle)
                raise OSError(error, "WaitForSingleObject failed")
            state = _read_state(self.state_path)
            if state.get("corrupt_state"):
                kernel32.ReleaseMutex(handle)
                kernel32.CloseHandle(handle)
                raise OfficeOperationQuarantined("Office lock state cannot be read; repair it before retry", {
                    "status": "OFFICE_OPERATION_STATE_UNREADABLE", "mutex": MUTEX_NAME,
                    "state_path": str(self.state_path),
                })
            quarantine = state.get("quarantine")
            if quarantine:
                kernel32.ReleaseMutex(handle)
                kernel32.CloseHandle(handle)
                raise OfficeOperationQuarantined("A timed-out native helper requires recovery before retry", {
                    "status": "OFFICE_OPERATION_QUARANTINED", "mutex": MUTEX_NAME,
                    "state_path": str(self.state_path), "quarantine": quarantine,
                    "recovery": "Call recover_quarantine after verifying the recorded child has exited.",
                })
            # A process can die while its helper remains alive.  Windows then
            # grants the abandoned mutex to us, so the state file is the only
            # durable record preventing concurrent Office access.
            active_child = state.get("active_child")
            if active_child:
                status, detail = _process_status(active_child)
                kernel32.ReleaseMutex(handle)
                kernel32.CloseHandle(handle)
                raise OfficeOperationQuarantined("Office lock has a child requiring recovery", {
                    "status": "OFFICE_OPERATION_ABANDONED_CHILD", "mutex": MUTEX_NAME,
                    "state_path": str(self.state_path), "child": active_child,
                    "child_status": status, "child_detail": detail,
                    "recovery": "Call recover_quarantine after the recorded child is proven exited.",
                })
            owner = _identity(self.operation, self.metadata)
            try:
                _write_state(self.state_path, {"owner": owner})
            except BaseException:
                kernel32.ReleaseMutex(handle)
                kernel32.CloseHandle(handle)
                raise
            _held = _Held(handle=handle, thread_id=current_thread, depth=1, owner=owner)
            self._outermost = True
            return self

    def register_child(self, pid: int) -> dict[str, Any]:
        """Durably record a child before waiting, including across owner death."""
        global _held
        with _local_guard:
            if _held is None or _held.thread_id != threading.get_ident():
                raise RuntimeError("register_child requires the current OfficeOperationLock owner")
            child = {"pid": pid, "creation_time": _creation_time(pid)}
            _write_state(self.state_path, {"owner": _held.owner, "active_child": child})
            return child

    def quarantine_child(self, pid: int, *, reason: str) -> dict[str, Any]:
        """Record a live helper that outlasted its timeout while mutex is owned."""
        global _held
        with _local_guard:
            if _held is None or _held.thread_id != threading.get_ident():
                raise RuntimeError("quarantine_child requires the current OfficeOperationLock owner")
            child = {"pid": pid, "creation_time": _creation_time(pid)}
            record = {"owner": _held.owner, "quarantine": {
                "child": child, "reason": reason, "recorded_at": _utc_now(),
                "status": "TIMED_OUT_CHILD_MAY_STILL_OPERATE_OFFICE",
            }}
            _write_state(self.state_path, record)
            return record["quarantine"]

    def __exit__(self, exc_type, exc, traceback) -> None:
        global _held
        with _local_guard:
            if _held is None or _held.thread_id != threading.get_ident():
                return None
            _held.depth -= 1
            if _held.depth:
                return None
            handle = _held.handle
            state = _read_state(self.state_path)
            # Preserve a timeout quarantine; ordinary successful/failed calls
            # remove their owner record once the native operation has finished.
            kernel32 = _kernel32()
            try:
                if state.get("quarantine"):
                    _write_state(self.state_path, {"quarantine": state["quarantine"]})
                else:
                    try:
                        self.state_path.unlink()
                    except FileNotFoundError:
                        pass
            finally:
                # State persistence failures must never strand the native mutex
                # or leave this process thinking it still owns it.
                _held = None
                try:
                    if not kernel32.ReleaseMutex(handle):
                        raise OSError(ctypes.get_last_error(), "ReleaseMutex failed")
                finally:
                    kernel32.CloseHandle(handle)
        return None


def recover_quarantine() -> dict[str, Any]:
    """Raw recovery implementation because ordinary acquisition blocks quarantine."""
    with _local_guard:
        if _held is not None:
            raise OfficeOperationBusy("Cannot recover while this process owns an Office operation", {
                "status": "OFFICE_OPERATION_BUSY", "mutex": MUTEX_NAME, "owner": _held.owner,
            })
    kernel32 = _kernel32()
    handle = kernel32.CreateMutexW(None, False, MUTEX_NAME)
    if not handle:
        raise OSError(ctypes.get_last_error(), "CreateMutexW failed")
    result = WAIT_TIMEOUT
    try:
        result = kernel32.WaitForSingleObject(handle, 0)
        if result == WAIT_TIMEOUT:
            raise OfficeOperationBusy("Another plugin native operation is active", {
                "status": "OFFICE_OPERATION_BUSY", "mutex": MUTEX_NAME, "state_path": str(_state_path()),
                "owner": _read_state(_state_path()).get("owner"),
            })
        if result not in (WAIT_OBJECT_0, WAIT_ABANDONED):
            raise OSError(ctypes.get_last_error(), "WaitForSingleObject failed")
        state = _read_state(_state_path())
        if state.get("corrupt_state"):
            raise OfficeOperationQuarantined("Office lock state cannot be read; recovery is unsafe", {
                "status": "OFFICE_OPERATION_STATE_UNREADABLE", "state_path": str(_state_path()),
            })
        quarantine = state.get("quarantine")
        active_child = state.get("active_child")
        if not quarantine and not active_child:
            return {"status": "NO_QUARANTINE", "state_path": str(_state_path())}
        child = quarantine.get("child", {}) if quarantine else active_child
        status, detail = _process_status(child)
        if status != "dead":
            raise OfficeOperationQuarantined("Recorded Office helper has not been proven exited; do not retry", {
                "status": "OFFICE_OPERATION_QUARANTINED", "state_path": str(_state_path()),
                "quarantine": quarantine, "active_child": active_child,
                "child_status": status, "child_detail": detail,
            })
        try:
            _state_path().unlink()
        except FileNotFoundError:
            pass
        return {"status": "QUARANTINE_CLEARED", "state_path": str(_state_path()),
                "quarantine": quarantine, "active_child": active_child}
    finally:
        # Waiting can fail before ownership; in that case ReleaseMutex would be
        # invalid.  WaitForSingleObject only reaches this finally after success
        # or a handled busy error, and the latter returns before this block.
        if result in (WAIT_OBJECT_0, WAIT_ABANDONED):
            kernel32.ReleaseMutex(handle)
        kernel32.CloseHandle(handle)


def run_locked_subprocess(command: Sequence[str | os.PathLike[str]], *, operation: str,
                          timeout_seconds: float, **popen_kwargs: Any) -> subprocess.CompletedProcess[str]:
    """Run one native helper under the lock and quarantine it on timeout.

    This function intentionally does not terminate the child.  A live child may
    still have an in-flight COM call.  Quarantine makes every later plugin entry
    fail fast until ``recover_quarantine`` proves that exact child instance exited.
    """
    if timeout_seconds <= 0:
        raise ValueError("timeout_seconds must be positive")
    with OfficeOperationLock(operation, metadata={"command": [str(x) for x in command]}) as lock:
        child = subprocess.Popen(command, **popen_kwargs)
        lock.register_child(child.pid)
        try:
            stdout, stderr = child.communicate(timeout=timeout_seconds)
        except subprocess.TimeoutExpired as error:
            quarantine = lock.quarantine_child(child.pid, reason=f"helper exceeded {timeout_seconds:g} seconds")
            raise OfficeOperationTimedOut("Native helper timed out and is quarantined; inspect/recover before retry", {
                "status": "OFFICE_OPERATION_TIMEOUT_QUARANTINED", "pid": child.pid,
                "state_path": str(lock.state_path), "quarantine": quarantine,
            }) from error
        return subprocess.CompletedProcess(command, child.returncode, stdout, stderr)


def main() -> None:
    """CLI wrapper for PowerShell entrypoints that launch one native helper.

    Example: ``python office_operation_lock.py --operation compose --timeout 180
    --command powershell.exe -File compose_slide.ps1 ...``.
    """
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--operation", default="office-operation")
    parser.add_argument("--timeout", type=float, default=180)
    parser.add_argument("--recover", action="store_true")
    parser.add_argument("--command", nargs=argparse.REMAINDER)
    args = parser.parse_args()
    try:
        if args.recover:
            if args.command:
                parser.error("--recover cannot be combined with --command")
            print(json.dumps(recover_quarantine()))
            return
        if not args.command:
            parser.error("--command is required unless --recover is used")
        completed = run_locked_subprocess(args.command, operation=args.operation,
                                          timeout_seconds=args.timeout)
        print(json.dumps({"status": "OFFICE_OPERATION_COMPLETE", "returncode": completed.returncode}))
        raise SystemExit(completed.returncode)
    except OfficeOperationLockError as error:
        print(json.dumps({"error": str(error), **error.diagnostic}), file=sys.stderr)
        raise SystemExit(75)


if __name__ == "__main__":
    main()
