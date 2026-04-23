"""Top-level entry point used by PyInstaller.

Uses absolute imports so PyInstaller can correctly trace the package's dependencies.
The package also has __main__.py for ``python -m incident_desk``."""
from incident_desk.app import App


if __name__ == "__main__":
    App().mainloop()
