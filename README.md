# Invitation Card Maker

A desktop app that stamps each guest's name onto your wedding invitation, one card per guest. Long names shrink automatically to fit.

Everything runs on your own computer. Nothing is uploaded anywhere.

## How to use

1. **Choose Image**: pick your invitation (JPEG or PNG).
2. **Download Excel Template**: save the template, fill in one guest (or couple) per row in column A, then **Choose Excel File** to load it.
3. **Font**: pick a style from the list. The preview updates straight away. Advanced users can also use their own `.ttf`/`.otf` file.
4. **Name position**: click the preview where the name should go. Fine-tune with the number boxes if needed. The red box is the area the name must fit inside; the blue line is the centre. Use **Preview with** to check long names. **Reset position** restores the defaults.
5. **Save to**: choose the output folder, then click **Generate Cards**.

### Your settings are remembered

The app saves your choices automatically: position, colour, font, output folder, file name prefix and window size. Your last invitation and guest list reopen automatically too. Settings are stored in your user profile, not in this folder, so they're never uploaded to GitHub:

- Windows: `%APPDATA%\Invitation Card Maker\settings.json`
- macOS: `~/Library/Application Support/Invitation Card Maker/settings.json`
- Linux: `~/.config/Invitation Card Maker/settings.json`

Delete that file to start fresh.

## Fonts

These fonts are built in. All are free under the [SIL Open Font License](fonts/licenses/).

| Style | Fonts |
| --- | --- |
| Script | Alex Brush, Great Vibes, Parisienne, Pinyon Script, Allura, Dancing Script |
| Classic | Cormorant Garamond, Playfair Display |
| Modern | Lato |
| Chinese (Windows only) | KaiTi 楷体, Microsoft YaHei 微软雅黑. These ship with Windows and are not bundled. |

Names written in Chinese automatically switch to KaiTi when an English-only font is selected, so they never show as empty boxes.

## Quick start (Windows)

1. Install [Python 3](https://www.python.org/downloads/) and tick **Add python.exe to PATH** during setup.
2. Double-click **`Invitation Card Maker.bat`**.

The first launch (and the first launch after an update) takes a minute while it installs what the app needs into a private `.venv` folder. After that, the app opens straight away. To reinstall from scratch, delete the `.venv` folder.

## Run from source

Requires Python 3.9 or newer, with Tkinter. Tkinter is included in the python.org installers. On Linux, install `python3-tk`.

```bash
pip install -r requirements.txt
python invitation_app.py
```

## Build a standalone app

To share with people who don't have Python installed:

```bash
pip install pyinstaller
python build.py
```

Send them the `dist/Invitation Card Maker/` folder. On Windows they double-click `Invitation Card Maker.exe`. PyInstaller builds for the operating system it runs on, so build on Windows for Windows and on macOS for macOS.

## Project layout

| File | Purpose |
| --- | --- |
| `Invitation Card Maker.bat` | Windows launcher: sets up `.venv` on first run, then starts the app |
| `invitation_app.py` | The desktop window (CustomTkinter) and saved settings |
| `cards.py` | Name fitting, Excel reading and template creation |
| `fonts/` | Bundled fonts, with their licences in `fonts/licenses/` |
| `build.py` | PyInstaller packaging |

## Privacy

Guest lists, generated cards, invitation artwork and licensed fonts are excluded in `.gitignore`.

## License

The code is released under the [MIT License](LICENSE).

The bundled fonts in `fonts/` are not covered by the MIT License. Each keeps its own [SIL Open Font License](fonts/licenses/), which allows free use and redistribution but not selling the fonts on their own. The Chinese fonts belong to Microsoft and come with Windows; they are not included in this repository.
