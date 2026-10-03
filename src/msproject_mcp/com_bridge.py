"""Thread-affine access to the MS Project COM server.

COM objects live in the apartment of the thread that created them, so every
call into Project is funnelled through one dedicated worker thread that has
called ``CoInitialize``. The async MCP tools submit work to that thread.
"""

from __future__ import annotations

import asyncio
import sys
import threading
from concurrent.futures import ThreadPoolExecutor
from typing import Any, Callable, TypeVar

T = TypeVar("T")

PROG_ID = "MSProject.Application"


class ProjectNotAvailableError(RuntimeError):
    """Raised when Microsoft Project cannot be reached over COM."""


def _com_thread_init() -> None:
    if sys.platform == "win32":
        import pythoncom

        pythoncom.CoInitialize()


def default_app_factory(visible: bool) -> Any:
    """Attach to a running Project instance or launch a new one."""
    if sys.platform != "win32":
        raise ProjectNotAvailableError(
            "Microsoft Project automation requires Windows with Project installed."
        )
    try:
        import win32com.client
    except ImportError as exc:  # pragma: no cover - Windows only
        raise ProjectNotAvailableError(
            "pywin32 is not installed. Run: pip install pywin32"
        ) from exc

    try:
        app = win32com.client.GetActiveObject(PROG_ID)
    except Exception:
        try:
            app = win32com.client.Dispatch(PROG_ID)
        except Exception as exc:  # pragma: no cover - Windows only
            raise ProjectNotAvailableError(
                f"Could not start Microsoft Project ({PROG_ID}): {exc}"
            ) from exc
    app.Visible = visible
    # Suppress modal dialogs that would otherwise block automation.
    app.DisplayAlerts = False
    return app


class ComBridge:
    """Owns the Project ``Application`` object and the thread it lives on."""

    def __init__(
        self,
        app_factory: Callable[[bool], Any] = default_app_factory,
        visible: bool = True,
    ) -> None:
        self._app_factory = app_factory
        self._visible = visible
        self._app: Any = None
        self._lock = threading.Lock()
        self._executor = ThreadPoolExecutor(
            max_workers=1,
            thread_name_prefix="msproject-com",
            initializer=_com_thread_init,
        )

    def app(self) -> Any:
        """Return the Application object. Must be called on the COM thread."""
        with self._lock:
            if self._app is not None and not self._is_alive(self._app):
                self._app = None
            if self._app is None:
                self._app = self._app_factory(self._visible)
            return self._app

    @staticmethod
    def _is_alive(app: Any) -> bool:
        try:
            _ = app.Name
            return True
        except Exception:
            return False

    def run_sync(self, fn: Callable[[Any], T]) -> T:
        """Run ``fn(app)`` on the COM thread and block for the result."""
        return self._executor.submit(lambda: fn(self.app())).result()

    async def run(self, fn: Callable[[Any], T]) -> T:
        """Run ``fn(app)`` on the COM thread without blocking the event loop."""
        loop = asyncio.get_running_loop()
        return await loop.run_in_executor(self._executor, lambda: fn(self.app()))

    def shutdown(self) -> None:
        self._executor.shutdown(wait=False)
