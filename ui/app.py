"""SafeSales order SMS — WinUI 3 shell (win32more).

This module is intentionally *import-safe* in headless environments (CI) where the
Windows App Runtime isn't installed.

The WinUI implementation lives in `ui._winui_app` and is imported lazily when
`main()` is called.
"""

from __future__ import annotations

from importlib import import_module
from typing import Any, Final

_IMPL_MODULE: Final[str] = "ui._winui_app"

_RUNTIME_MISSING_MESSAGE: Final[str] = (
    "WinUI runtime not found. This application requires the Windows App Runtime "
    "Version 2.0 (MSIX package version >= 2.0.1.0). "
    "This dependency cannot be bundled into the one-file EXE, so it must be "
    "installed separately.\n\n"
    "Download and install the Windows App Runtime from: "
    "https://learn.microsoft.com/windows/apps/windows-app-sdk/prepare-systems"
)


def _looks_like_runtime_missing(exc: BaseException) -> bool:
    name = exc.__class__.__name__
    mod = getattr(exc.__class__, "__module__", "")
    msg = str(exc)
    if name == "RuntimeNotFoundError":
        return True
    if "Windows App Runtime" in msg and "Runtime" in name:
        return True
    if mod.startswith("win32more.appsdk") and "Windows App Runtime" in msg:
        return True
    return False


def main() -> None:
    """Start the WinUI app.

    Importing the WinUI bindings requires the Windows App Runtime. If it's not
    installed, raise a clear error message.
    """

    try:
        impl = import_module(_IMPL_MODULE)
    except Exception as exc:  # noqa: BLE001
        if _looks_like_runtime_missing(exc):
            raise RuntimeError(_RUNTIME_MISSING_MESSAGE) from exc
        raise
    impl.main()


def __getattr__(name: str) -> Any:
    """Lazy attribute access for the underlying WinUI implementation.

    This keeps `import ui.app` working in CI, while still allowing power-users to
    access implementation symbols (e.g. `SafeSalesWinUIApp`) when the runtime is
    available.
    """

    impl = import_module(_IMPL_MODULE)
    return getattr(impl, name)


if __name__ == "__main__":
    main()
