"""Entry point: ``python -m incident_desk`` and the PyInstaller target."""
from .app import App


def main() -> None:
    App().mainloop()


if __name__ == "__main__":
    main()
