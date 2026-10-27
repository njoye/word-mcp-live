"""COM connection manager for Microsoft Word on Windows.

Provides functions to connect to a running Word instance and find open documents.
Only works on Windows with pywin32 installed.
"""

import os
import sys
import unicodedata
from contextlib import contextmanager


def _get_app_from_file(filename: str):
    """Attempt to reach Word.Application by binding to an open document via GetObject."""
    if not filename:
        return None
    try:
        import win32com.client
        candidates = []
        if os.path.isabs(filename):
            candidates.append(filename)
        else:
            user_docs = os.path.join(os.path.expanduser('~'), 'Documents', filename)
            candidates.append(user_docs)
            candidates.append(os.path.abspath(filename))
            candidates.append(filename)

        for path in candidates:
            if os.path.exists(path):
                doc_obj = win32com.client.GetObject(os.path.abspath(path))
                if doc_obj is not None and hasattr(doc_obj, "Application"):
                    app = doc_obj.Application
                    # Keep COM reference alive on the application instance so Word doesn't release the document
                    app._doc_ref = doc_obj
                    if hasattr(app, "Documents") and app.Documents.Count > 0:
                        return app
    except Exception:
        pass
    return None


_CACHED_APP = None


def get_word_app(filename: str = None):
    """Get a reference to the running Word application via COM.

    Returns the Word.Application COM object that has open documents.
    When multiple Word instances are running, iterates through all
    Running Object Table (ROT) entries to find one with documents.
    Raises RuntimeError if Word is not running or not on Windows.
    """
    if sys.platform != "win32":
        raise RuntimeError("Word COM automation is only available on Windows")

    global _CACHED_APP
    if _CACHED_APP is not None:
        try:
            if hasattr(_CACHED_APP, "Documents") and _CACHED_APP.Documents.Count > 0:
                return _CACHED_APP
        except Exception:
            _CACHED_APP = None

    import win32com.client

    try:
        app = win32com.client.GetActiveObject("Word.Application")
        if app.Documents.Count > 0:
            _CACHED_APP = app
            return app
        # GetActiveObject found an empty instance — scan ROT for others
        app_with_docs = _find_word_with_docs()
        if app_with_docs is not None:
            _CACHED_APP = app_with_docs
            return app_with_docs
        if filename:
            app_from_file = _get_app_from_file(filename)
            if app_from_file is not None:
                _CACHED_APP = app_from_file
                return app_from_file
        # No instance has documents; return the empty one (caller may open a file)
        return app
    except Exception:
        # GetActiveObject failed entirely — try ROT scan
        app_with_docs = _find_word_with_docs()
        if app_with_docs is not None:
            _CACHED_APP = app_with_docs
            return app_with_docs
        if filename:
            app_from_file = _get_app_from_file(filename)
            if app_from_file is not None:
                _CACHED_APP = app_from_file
                return app_from_file
        raise RuntimeError(
            "Microsoft Word is not running. Please open Word first."
        )


def _find_word_with_docs():
    """Scan the Running Object Table for a Word.Application with open docs.

    Handles Office 365 / OneDrive scenarios where GetActiveObject returns an
    empty Application proxy.  In these cases, documents are registered in the
    ROT as file monikers (.docx paths or https://d.docs.live.net/... URLs).
    We grab the Document COM object from such a moniker and reach the real
    Application via ``doc.Application``.

    Returns the Word.Application COM object if found, or None.
    """
    try:
        import pythoncom
        import win32com.client

        rot = pythoncom.GetRunningObjectTable(0)
        enum = rot.EnumRunning()

        # Pass 1: look for a Word.Application ROT entry with documents
        monikers_to_retry = []
        while True:
            batch = enum.Next(1)
            if not batch:
                break
            moniker = batch[0]
            try:
                ctx = pythoncom.CreateBindCtx(0)
                name = moniker.GetDisplayName(ctx, None)
                obj = rot.GetObject(moniker)
                dispatch = obj.QueryInterface(pythoncom.IID_IDispatch)
                com_obj = win32com.client.Dispatch(dispatch)
                # Direct Application entry
                if hasattr(com_obj, "Documents") and hasattr(com_obj, "ActiveDocument"):
                    if com_obj.Documents.Count > 0:
                        return com_obj
                # Remember file monikers for pass 2
                if name and (name.lower().endswith(".docx") or name.lower().endswith(".doc")):
                    monikers_to_retry.append((name, moniker))
            except Exception:
                # Also collect file monikers we couldn't QI yet
                try:
                    ctx = pythoncom.CreateBindCtx(0)
                    name = moniker.GetDisplayName(ctx, None)
                    if name and (name.lower().endswith(".docx") or name.lower().endswith(".doc")):
                        monikers_to_retry.append((name, moniker))
                except Exception:
                    pass
                continue

        # Pass 2: try file monikers → Document → Application
        for name, moniker in monikers_to_retry:
            try:
                obj = rot.GetObject(moniker)
                dispatch = obj.QueryInterface(pythoncom.IID_IDispatch)
                doc = win32com.client.Dispatch(dispatch)
                app = doc.Application
                if app.Documents.Count > 0:
                    return app
            except Exception:
                continue
    except Exception:
        pass
    return None


def find_document(app, filename: str = None):
    """Find an open document by filename.

    Args:
        app: Word.Application COM object.
        filename: Document name (basename) or full path.
                  If None or empty, returns the active document.

    Returns:
        Document COM object.

    Raises:
        ValueError: If the document is not found or no documents are open.
    """
    if app.Documents.Count == 0:
        raise ValueError("No documents are open in Word")

    if not filename:
        active_doc = app.ActiveDocument
        if app.Documents.Count > 1:
            try:
                print(
                    f"[Word COM Warning] Multiple documents open ({app.Documents.Count}). "
                    f"Targeting active document '{active_doc.Name}'. "
                    f"Specify 'filename' explicitly to prevent accidental mis-targeting.",
                    file=sys.stderr,
                )
            except Exception:
                pass
        return active_doc

    target_basename = unicodedata.normalize('NFC', os.path.basename(filename)).lower()
    target_fullpath = (
        unicodedata.normalize('NFC', os.path.normpath(filename)).lower()
        if os.path.isabs(filename) else None
    )

    for i in range(1, app.Documents.Count + 1):
        doc = app.Documents(i)
        if unicodedata.normalize('NFC', doc.Name).lower() == target_basename:
            return doc
        if target_fullpath and unicodedata.normalize('NFC', os.path.normpath(doc.FullName)).lower() == target_fullpath:
            return doc

    open_docs = [app.Documents(i).Name for i in range(1, app.Documents.Count + 1)]
    raise ValueError(
        f"Document '{filename}' is not open in Word. "
        f"Open documents: {open_docs}"
    )


@contextmanager
def undo_record(app, name: str):
    """Wrap a block of COM mutations in a single Word UndoRecord.

    Groups all changes into one Ctrl+Z entry in Word's undo stack.
    The undo record name appears in Edit > Undo and in the undo history.
    Degrades gracefully on Word 2007 or earlier (no UndoRecord support).

    Args:
        app: Word.Application COM object.
        name: Label for the undo entry (truncated to 64 chars by Word).

    Usage::

        with undo_record(app, "MCP: Insert Text"):
            doc.Range(0, 0).InsertBefore("Hello")
    """
    rec = None
    try:
        rec = app.UndoRecord
        # Clean up stale undo record from a previous crash/interrupted session
        if rec.IsRecordingCustomRecord:
            try:
                rec.EndCustomRecord()
            except Exception:
                pass
        rec.StartCustomRecord(name[:64])
    except Exception:
        rec = None  # Word 2007 or earlier — proceed without
    try:
        yield
    finally:
        if rec is not None:
            try:
                rec.EndCustomRecord()
            except Exception:
                pass


def get_document_target_warning(app, filename: str = None) -> str:
    """Return a warning string if filename was omitted while multiple documents are open."""
    if not filename:
        try:
            count = app.Documents.Count
            if count > 1:
                active_name = app.ActiveDocument.Name
                return (
                    f"Multiple documents open in Word ({count}). 'filename' was not specified; "
                    f"targeted active document '{active_name}'. Pass 'filename' explicitly to prevent targeting the wrong document."
                )
        except Exception:
            pass
    return None


class AsyncRLock:
    """Re-entrant asyncio lock for serializing Word COM calls across Word's Single-Threaded Apartment (STA)."""

    def __init__(self):
        self._lock = None
        self._owner = None
        self._count = 0

    def _get_lock(self):
        if self._lock is None:
            import asyncio
            self._lock = asyncio.Lock()
        return self._lock

    async def acquire(self):
        import asyncio
        me = asyncio.current_task()
        if self._owner == me:
            self._count += 1
            return True
        await self._get_lock().acquire()
        self._owner = me
        self._count = 1
        return True

    def release(self):
        import asyncio
        me = asyncio.current_task()
        if self._owner != me:
            raise RuntimeError("Cannot release un-acquired lock")
        self._count -= 1
        if self._count == 0:
            self._owner = None
            self._get_lock().release()

    async def __aenter__(self):
        await self.acquire()
        return self

    async def __aexit__(self, exc_type, exc, tb):
        self.release()


_WORD_COM_LOCK = AsyncRLock()


def serialize_word_com(func):
    """Decorator to serialize async COM tool calls across Word's Single-Threaded Apartment (STA)
    and retry transient RPC_E_CALL_REJECTED (-2147418111) errors."""
    import functools
    import asyncio
    import inspect

    @functools.wraps(func)
    async def wrapper(*args, **kwargs):
        async with _WORD_COM_LOCK:
            max_retries = 3
            for attempt in range(max_retries):
                try:
                    if inspect.iscoroutinefunction(func):
                        return await func(*args, **kwargs)
                    res = func(*args, **kwargs)
                    if asyncio.iscoroutine(res):
                        return await res
                    return res
                except Exception as e:
                    err_str = str(e)
                    is_rejected = (
                        "-2147418111" in err_str
                        or "0x80010001" in err_str
                        or "rejected by callee" in err_str.lower()
                    )
                    if is_rejected and attempt < max_retries - 1:
                        await asyncio.sleep(0.15 * (attempt + 1))
                        continue
                    raise

    return wrapper

