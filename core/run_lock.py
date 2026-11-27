"""Crash-recoverable claims for local audit-flow output directories."""

from contextlib import contextmanager
import errno
import json
import os
import re
import socket
import stat
import sys
from time import time_ns
from uuid import uuid4

from core.schema import AuditError

_OWNER_FILE = re.compile(r"owner-([0-9a-f]{32})\.json")
_METADATA_LIMIT = 2048


def _unknown(claim):
    return AuditError(f"无法可靠确认运行锁状态：{claim}。请在本机核对任务及锁持有者，确认相关任务均已停止后，"
                      "手动移走此锁目录，再重试。")


def _try_lock(stream):
    """Hold an OS lock for the lifetime of the open descriptor."""
    try:
        if os.name == "nt":
            import msvcrt

            # The byte outside the metadata remains lockable even past EOF.
            stream.seek(_METADATA_LIMIT)
            msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
        else:
            import fcntl

            fcntl.flock(stream.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
    except OSError as error:
        if error.errno in {errno.EACCES, errno.EAGAIN, errno.EWOULDBLOCK, errno.EDEADLK}:
            return False
        raise
    return True


def _checked_stat(path, is_expected_type):
    entry = path.lstat()
    if not is_expected_type(entry.st_mode) or getattr(entry, "st_file_attributes", 0) & stat.FILE_ATTRIBUTE_REPARSE_POINT:
        raise ValueError
    return entry


def _read_owner(stream, token):
    entry = os.fstat(stream.fileno())
    if not stat.S_ISREG(entry.st_mode) or entry.st_nlink != 1 or entry.st_size > _METADATA_LIMIT:
        raise ValueError
    stream.seek(0)
    data = json.loads(stream.read(_METADATA_LIMIT))
    if (not isinstance(data, dict) or set(data) != {"version", "pid", "created_at_ns", "token", "hostname"}
            or type(data["version"]) is not int or data["version"] != 1
            or type(data["pid"]) is not int or data["pid"] <= 0
            or type(data["created_at_ns"]) is not int or data["created_at_ns"] <= 0
            or data["token"] != token or not isinstance(data["hostname"], str) or not data["hostname"]):
        raise ValueError
    return data


def _remove_owner(claim, marker, identity):
    _checked_stat(claim, stat.S_ISDIR)
    if tuple(claim.iterdir()) != (marker,) or not os.path.samestat(_checked_stat(marker, stat.S_ISREG), identity):
        return False
    # Only the process that unlinks this exact token can remove the directory.
    # A second reclaimer cannot remove a newly acquired directory by racing rmdir.
    marker.unlink()
    claim.rmdir()
    return True


def _reclaim(claim, busy):
    try:
        _checked_stat(claim, stat.S_ISDIR)
        entries = tuple(claim.iterdir())
        if len(entries) != 1 or (name := _OWNER_FILE.fullmatch(entries[0].name)) is None:
            raise ValueError
        marker, token = entries[0], name[1]
        identity = _checked_stat(marker, stat.S_ISREG)
        with marker.open("r+b", buffering=0) as stream:
            if not os.path.samestat(identity, os.fstat(stream.fileno())):
                raise ValueError
            owner = _read_owner(stream, token)
            if owner["hostname"] != socket.gethostname():
                raise ValueError
            if not _try_lock(stream):
                raise busy
            if _read_owner(stream, token) != owner:
                raise ValueError
        # The OS releases the descriptor lock on process exit. PID reuse and
        # wall-clock timestamps therefore play no part in deciding ownership.
        if not _remove_owner(claim, marker, identity):
            raise ValueError
    except FileNotFoundError:
        if os.path.lexists(claim):
            raise _unknown(claim) from None
    except (OSError, ValueError, RecursionError, ImportError):
        raise _unknown(claim) from None


@contextmanager
def run_lock(claim, busy_message):
    busy = AuditError(f"{busy_message} 锁位置：{claim}。请等待任务结束。")
    token = uuid4().hex
    owner = {"version": 1, "pid": os.getpid(), "created_at_ns": time_ns(), "token": token,
             "hostname": socket.gethostname()}
    try:
        try:
            claim.mkdir()
        except FileExistsError:
            _reclaim(claim, busy)
            claim.mkdir()
    except FileExistsError:
        raise busy from None
    except OSError:
        raise _unknown(claim) from None

    marker = claim / f"owner-{token}.json"
    identity = None
    keep_ownership = True
    try:
        try:
            stream = marker.open("x+b", buffering=0)
        except OSError:
            raise _unknown(claim) from None
        with stream:
            try:
                identity = os.fstat(stream.fileno())
                if not _try_lock(stream):
                    raise _unknown(claim)
                stream.seek(0)
                payload = json.dumps(owner).encode("utf-8")
                if stream.write(payload) != len(payload):
                    raise OSError
            except (OSError, ImportError):
                raise _unknown(claim) from None
            try:
                yield
            finally:
                try:
                    keep_ownership = _read_owner(stream, token) == owner
                except (OSError, ValueError, RecursionError):
                    keep_ownership = False
    finally:
        if identity is not None and keep_ownership:
            pending_error = sys.exc_info()[0] is not None
            try:
                _remove_owner(claim, marker, identity)
            except (FileNotFoundError, ValueError):
                pass  # The claim has already disappeared or no longer matches this owner.
            except OSError:
                if not pending_error:
                    raise _unknown(claim) from None
