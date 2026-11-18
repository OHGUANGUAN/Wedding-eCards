#!/bin/bash
# Double-click to start Invitation Card Maker on macOS.
# The first run (and the first run after an update) installs what the app needs
# into a private Python environment in .venv. This takes a minute.
cd "$(dirname "$0")" || exit 1

VENV=.venv
MARKER="$VENV/requirements.installed"

fail() {
    echo
    echo "$1"
    echo
    read -r -n 1 -p "Press any key to close."
    exit 1
}

# Skip setup when the installed requirements match requirements.txt
if ! cmp -s requirements.txt "$MARKER"; then
    echo "Setting up Invitation Card Maker..."
    echo
    if [ ! -x "$VENV/bin/python" ]; then
        command -v python3 >/dev/null && python3 -c "import tkinter" 2>/dev/null ||
            fail "Python 3 was not found. Install it from https://www.python.org/downloads/ then run this file again."
        python3 -m venv "$VENV" || fail "Could not create the Python environment."
    fi
    "$VENV/bin/python" -m pip install --upgrade pip --quiet --disable-pip-version-check
    "$VENV/bin/python" -m pip install -r requirements.txt --quiet --disable-pip-version-check ||
        fail "Setup failed. Check your internet connection and try again. To start over, delete the .venv folder."
    cp requirements.txt "$MARKER"
    echo "Setup complete."
fi

nohup "$VENV/bin/python" invitation_app.py >/dev/null 2>&1 &
