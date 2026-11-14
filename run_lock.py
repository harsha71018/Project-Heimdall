"""Process-owned run locks. The OS releases a lock if its owner crashes."""
import os

if os.name == "nt":
    import msvcrt
else:
    import fcntl

_owned_locks = {}


def acquire_run_lock(lock_file="heimdall.lock"):
    path = os.path.normcase(os.path.abspath(lock_file))
    if path in _owned_locks:
        return False
    handle = None
    try:
        handle = open(path, "a+b")
        if os.path.getsize(path) == 0:
            handle.write(b"\0")
            handle.flush()
        handle.seek(0)
        if os.name == "nt":
            msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
        else:
            fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        _owned_locks[path] = handle
        return True
    except OSError:
        if handle is not None:
            handle.close()
        return False


def release_run_lock(lock_file="heimdall.lock"):
    path = os.path.normcase(os.path.abspath(lock_file))
    handle = _owned_locks.pop(path, None)
    if handle is None:
        return
    try:
        handle.seek(0)
        if os.name == "nt":
            msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
        else:
            fcntl.flock(handle.fileno(), fcntl.LOCK_UN)
    finally:
        handle.close()
    # Do not unlink: a waiter may already have opened the same file.
