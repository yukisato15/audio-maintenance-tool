from __future__ import annotations

from pathlib import Path
import subprocess
import sys


ROOT = Path(__file__).resolve().parent


def main() -> None:
    subprocess.run([sys.executable, "generate_app_icon.py"], cwd=ROOT, check=True)

    import tkinterdnd2  # noqa: PLC0415

    tkdnd_dir = Path(tkinterdnd2.__file__).resolve().parent / "tkdnd"
    command = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--windowed",
        "--name",
        "AudioMaintenanceTool",
        "--icon",
        "assets/audio_maintenance_tool.ico",
        "--collect-data",
        "customtkinter",
        "--hidden-import",
        "tkinterdnd2",
        "--add-data",
        f"{tkdnd_dir};tkinterdnd2/tkdnd",
        "main.py",
    ]
    subprocess.run(command, cwd=ROOT, check=True)


if __name__ == "__main__":
    main()
