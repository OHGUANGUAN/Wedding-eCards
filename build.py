"""Package the app as a standalone program with PyInstaller.

    pip install pyinstaller
    python build.py

The result is in dist/Invitation Card Maker/ (on macOS: dist/Invitation Card Maker.app).
"""

import os

import PyInstaller.__main__

PyInstaller.__main__.run([
    "invitation_app.py",
    "--name", "Invitation Card Maker",
    "--windowed",
    "--noconfirm",
    "--add-data", f"fonts{os.pathsep}fonts",
    "--collect-data", "customtkinter",  # its theme files
])
