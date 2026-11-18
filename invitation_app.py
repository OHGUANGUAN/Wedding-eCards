"""Desktop app: stamp each guest's name from an Excel list onto an invitation image."""

import json
import os
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import colorchooser, filedialog, messagebox

import customtkinter as ctk
from PIL import Image, ImageDraw, ImageFont, ImageOps, ImageTk

import cards

APP_NAME = "Invitation Card Maker"
IS_WINDOWS, IS_MAC = sys.platform == "win32", sys.platform == "darwin"
PLACEMENT_FIELDS = {
    "x_center": "Centre X", "base_y": "Top Y", "x_min": "Left limit",
    "max_y": "Max Y", "x_max": "Right limit", "max_font_size": "Max font size",
}
PREVIEW_PLACEHOLDER = "Guest Name"
EXCEL_TYPES = [("Excel workbook", "*.xlsx")]

# --- Palette ---
BG, CARD, FIELD, BORDER = "#F7F1EC", "#FFFFFF", "#FBF8F6", "#EADFD6"
ACCENT, ACCENT_HOVER, ACCENT_SOFT = "#B5735C", "#9C5E49", "#F3E4DC"
TEXT, MUTED, PREVIEW_BG = "#2F2A27", "#8A7F78", "#EDE6E0"


# --- Settings (saved in the user's profile, never in the project folder) ---

def settings_path():
    if IS_WINDOWS:
        base = Path(os.environ.get("APPDATA", Path.home()))
    elif IS_MAC:
        base = Path.home() / "Library" / "Application Support"
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config"))
    return base / APP_NAME / "settings.json"


def load_settings():
    try:
        return json.loads(settings_path().read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def save_settings(data):
    path = settings_path()
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix(".tmp")
        tmp.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
        tmp.replace(path)
    except OSError:
        pass  # settings are a convenience; never block the app


def open_folder(path):
    if IS_WINDOWS:
        os.startfile(path)
    else:
        subprocess.run(["open" if IS_MAC else "xdg-open", str(path)])


def plural(n, word):
    return f"{n} {word}{'' if n == 1 else 's'}"


# --- Styled widgets ---

def font(size=13, weight="normal"):
    return ctk.CTkFont(size=size, weight=weight)


BUTTON_STYLES = {
    "primary": lambda: dict(fg_color=ACCENT, hover_color=ACCENT_HOVER, text_color="white", font=font(13, "bold")),
    "secondary": lambda: dict(fg_color="transparent", hover_color=ACCENT_SOFT, border_width=1, border_color=ACCENT,
                              text_color=ACCENT, font=font(13)),
    "link": lambda: dict(fg_color="transparent", hover_color=ACCENT_SOFT, text_color=ACCENT, height=26, width=0,
                         anchor="w", font=font(12, "bold")),
}


def button(parent, text, command, style="primary", **kw):
    return ctk.CTkButton(parent, text=text, command=command,
                         **{"height": 34, "corner_radius": 8, **BUTTON_STYLES[style](), **kw})


def label(parent, text="", size=12, weight="normal", color=TEXT, **kw):
    return ctk.CTkLabel(parent, text=text, text_color=color, font=font(size, weight), **kw)


def muted_label(parent, wraplength=330, **kw):
    return label(parent, color=MUTED, anchor="w", justify="left", wraplength=wraplength, **kw)


def row(parent, **pack):
    frame = ctk.CTkFrame(parent, fg_color="transparent")
    frame.pack(fill="x", **pack)
    return frame


def entry(parent, var, **kw):
    return ctk.CTkEntry(parent, textvariable=var, fg_color=FIELD, border_color=BORDER, text_color=TEXT,
                        corner_radius=8, **kw)


def option_menu(parent, var, values):
    return ctk.CTkOptionMenu(
        parent, variable=var, values=values, fg_color=FIELD, button_color=BORDER, button_hover_color=ACCENT_SOFT,
        text_color=TEXT, dropdown_fg_color=CARD, dropdown_hover_color=ACCENT_SOFT, dropdown_text_color=TEXT,
        corner_radius=8, dynamic_resizing=False, anchor="w", font=font(13),
    )


def step_card(parent, number, title):
    """A white rounded card with a numbered heading; returns the body frame."""
    card = ctk.CTkFrame(parent, fg_color=CARD, border_color=BORDER, border_width=1, corner_radius=14)
    card.pack(fill="x", pady=(0, 12), padx=(0, 6))
    header = row(card, padx=16, pady=(14, 8))
    ctk.CTkLabel(header, text=str(number), width=26, height=26, corner_radius=13,
                 fg_color=ACCENT, text_color="white", font=font(13, "bold")).pack(side="left")
    label(header, title, 15, "bold").pack(side="left", padx=(10, 0))
    return row(card, padx=16, pady=(0, 14))


class App(ctk.CTk):
    def __init__(self):
        super().__init__(fg_color=BG)
        self.title(APP_NAME)
        self.minsize(820, 520)

        self.settings = load_settings()
        self.base_image = self.image_path = self.names_path = self.custom_font_path = None
        self.names = []
        self.fonts = {f.label: f for f in cards.available_fonts()}
        self.color = cards.Placement.color
        self.last_dir = self.settings.get("last_dir") or str(Path.home())
        self.preview_photo = None
        self.preview_scale = 1.0
        self._preview_job = self._save_job = None
        self._jobs = []

        self.placement_vars = {field: tk.StringVar() for field in PLACEMENT_FIELDS}
        self.image_text = tk.StringVar(value="No image chosen yet")
        self.names_text = tk.StringVar(value="No guest list chosen yet")
        self.font_label = tk.StringVar(value=next(iter(self.fonts)))
        self.font_hint = tk.StringVar()
        self.color_text = tk.StringVar()
        self.preview_name = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.prefix = tk.StringVar(value="Invite_")
        self.status = tk.StringVar(value="Choose an invitation image and a guest list to begin.")

        self._build_sidebar()
        self._build_preview()
        self.set_placement(cards.Placement())
        self._restore_window()
        self._restore_settings()

        for var in [*self.placement_vars.values(), self.preview_name]:
            var.trace_add("write", lambda *_: self.schedule_preview())
        self.font_label.trace_add("write", lambda *_: (self.update_font_hint(), self.schedule_preview()))
        for var in [*self.placement_vars.values(), self.font_label, self.output_dir, self.prefix]:
            var.trace_add("write", lambda *_: self.schedule_save())
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    # --- Layout ---

    def _build_sidebar(self):
        sidebar = ctk.CTkFrame(self, fg_color=BG, corner_radius=0)
        sidebar.pack(side="left", fill="y", padx=(16, 8), pady=16)

        header = row(sidebar, pady=(0, 12))
        label(header, APP_NAME, 22, "bold").pack(anchor="w")
        label(header, "A personalised wedding card for every guest", 13, color=MUTED).pack(anchor="w")

        # Packed before the scrolling steps so it always stays visible at the bottom
        footer = row(sidebar, side="bottom", pady=(12, 0))
        self.generate_btn = button(footer, "Generate Cards", self.generate, height=44, font=font(15, "bold"))
        self.generate_btn.pack(fill="x")
        self.progress = ctk.CTkProgressBar(footer, progress_color=ACCENT, fg_color=BORDER, height=6)
        self.progress.set(0)
        self.progress.pack(fill="x", pady=(10, 6))
        muted_label(footer, textvariable=self.status, wraplength=360).pack(fill="x")

        steps = ctk.CTkScrollableFrame(
            sidebar, fg_color=BG, width=370, scrollbar_button_color=BORDER, scrollbar_button_hover_color=ACCENT_SOFT
        )
        steps.pack(fill="both", expand=True)

        body = step_card(steps, 1, "Invitation image")
        button(body, "Choose Image", self.choose_image).pack(anchor="w")
        muted_label(body, textvariable=self.image_text).pack(fill="x", pady=(6, 0))

        body = step_card(steps, 2, "Guest list")
        r = row(body)
        button(r, "Choose Excel File", self.choose_names).pack(side="left")
        button(r, "Download Excel Template", self.save_template, "secondary").pack(side="left", padx=(8, 0))
        muted_label(body, textvariable=self.names_text).pack(fill="x", pady=(6, 0))

        body = step_card(steps, 3, "Font")
        self.font_menu = option_menu(body, self.font_label, list(self.fonts))
        self.font_menu.pack(fill="x")
        button(body, "Use my own font file…", self.choose_font, "link").pack(anchor="w", pady=(4, 0))
        muted_label(body, textvariable=self.font_hint).pack(fill="x")

        body = step_card(steps, 4, "Name position")
        muted_label(body, text="Click the preview to place the name, or fine-tune in image pixels.").pack(fill="x")
        grid = row(body, pady=(8, 0))
        for i, (field, text) in enumerate(PLACEMENT_FIELDS.items()):
            r, c = divmod(i, 2)
            label(grid, text).grid(row=r, column=c * 2, sticky="w", pady=3)
            entry(grid, self.placement_vars[field], width=72).grid(row=r, column=c * 2 + 1, sticky="w",
                                                                   padx=(6, 14), pady=3)
        r = row(body, pady=(8, 0))
        label(r, "Colour").pack(side="left")
        self.color_btn = ctk.CTkButton(r, text="", width=30, height=26, corner_radius=6, border_width=1,
                                       border_color=BORDER, command=self.choose_color)
        self.color_btn.pack(side="left", padx=(8, 6))
        label(r, textvariable=self.color_text, color=MUTED).pack(side="left")
        button(r, "Reset position", self.reset_placement, "link").pack(side="right")
        label(body, "Preview with").pack(anchor="w", pady=(10, 2))
        self.preview_menu = option_menu(body, self.preview_name, [PREVIEW_PLACEHOLDER])
        self.preview_menu.pack(fill="x")

        body = step_card(steps, 5, "Save to")
        r = row(body)
        entry(r, self.output_dir).pack(side="left", fill="x", expand=True)
        button(r, "Browse", self.choose_output, "secondary", width=80).pack(side="left", padx=(8, 0))
        r = row(body, pady=(8, 0))
        label(r, "File name prefix").pack(side="left")
        entry(r, self.prefix, width=140).pack(side="left", padx=(8, 0))

    def _build_preview(self):
        pane = ctk.CTkFrame(self, fg_color=CARD, border_color=BORDER, border_width=1, corner_radius=14)
        pane.pack(side="left", fill="both", expand=True, padx=(8, 16), pady=16)
        header = row(pane, padx=18, pady=(14, 8))
        label(header, "Preview", 15, "bold").pack(side="left")
        label(header, "Red box: space the name must fit in.  Blue line: centre.", color=MUTED).pack(side="right")

        area = ctk.CTkFrame(pane, fg_color="transparent")
        area.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        self.canvas = tk.Canvas(area, background=PREVIEW_BG, highlightthickness=0, cursor="crosshair")
        scroll = ctk.CTkScrollbar(area, command=self.canvas.yview, button_color=BORDER, button_hover_color=ACCENT_SOFT)
        self.canvas.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y", padx=(6, 0))
        self.canvas.pack(side="left", fill="both", expand=True)

        self.canvas.bind("<Configure>", lambda _: self.schedule_preview())
        self.canvas.bind("<Button-1>", self.on_preview_click)
        # Windows/macOS send <MouseWheel> (delta is ±120 on Windows, small values on macOS); Linux sends Button-4/5
        for event in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.canvas.bind(event, self.on_scroll)

    def on_scroll(self, event):
        self.canvas.yview_scroll(-1 if event.num == 4 or event.delta > 0 else 1, "units")

    # --- Settings ---

    def _restore_window(self):
        self.geometry(self.settings.get("geometry") or "1200x800")
        # "zoomed" (maximised) is a Windows-only window state in Tk
        if IS_WINDOWS and self.settings.get("zoomed", True):
            self.after(0, lambda: self.state("zoomed"))

    def _restore_settings(self):
        s = self.settings
        for field, value in s.get("placement", {}).items():
            if field in self.placement_vars:
                self.placement_vars[field].set(str(value))
        if s.get("color"):
            self._set_color(s["color"])
        if s.get("custom_font") and Path(s["custom_font"]).exists():
            self._add_custom_font(Path(s["custom_font"]))
        if s.get("font") in self.fonts:
            self.font_label.set(s["font"])
        self.output_dir.set(s.get("output_dir", ""))
        self.prefix.set(s.get("prefix", "Invite_"))

        reopened = [what for what, key, load in [("invitation", "image", self.load_image),
                                                 ("guest list", "names", self.load_names)]
                    if s.get(key) and Path(s[key]).exists() and load(s[key], quiet=True)]
        self.update_font_hint()
        if reopened:
            verb = "were" if len(reopened) > 1 else "was"
            self.status.set(f"Welcome back! Your last {' and '.join(reopened)} {verb} reopened.")

    def schedule_save(self):
        if self._save_job:
            self.after_cancel(self._save_job)
        self._save_job = self.after(800, self.save_now)

    def save_now(self):
        self._save_job = None
        zoomed = self.state() == "zoomed"
        if not zoomed:  # keep the last normal-size geometry, so un-maximising restores it
            self.settings["geometry"] = self.geometry()
        save_settings({
            "placement": {field: var.get() for field, var in self.placement_vars.items()},
            "color": self.color,
            "font": self.font_label.get(),
            "custom_font": str(self.custom_font_path) if self.custom_font_path else None,
            "output_dir": self.output_dir.get(),
            "prefix": self.prefix.get(),
            "image": self.image_path,
            "names": self.names_path,
            "last_dir": self.last_dir,
            "zoomed": zoomed,
            "geometry": self.settings.get("geometry"),
        })

    def on_close(self):
        self.save_now()
        self.destroy()

    # --- Inputs ---

    def _ask_open(self, **kw):
        path = filedialog.askopenfilename(initialdir=self.last_dir, **kw)
        if path:
            self.last_dir = str(Path(path).parent)
        return path

    def choose_image(self):
        path = self._ask_open(
            title="Choose invitation image", filetypes=[("Images", "*.jpg *.jpeg *.png"), ("All files", "*.*")]
        )
        if path and self.load_image(path):
            self.update_ready_status()

    def load_image(self, path, quiet=False):
        try:
            self.base_image = ImageOps.exif_transpose(Image.open(path)).convert("RGB")
        except Exception as e:
            if not quiet:
                messagebox.showerror("Could not open image", str(e))
            return False
        path = Path(path).resolve()
        self.image_path = str(path)
        w, h = self.base_image.size
        self.image_text.set(f"{path.name}  ·  {w} × {h} px")
        if not self.output_dir.get():
            self.output_dir.set(str(path.parent / "personalized_cards"))
        self.render_preview()
        self.after(50, self.scroll_to_name)
        self.schedule_save()
        return True

    def choose_names(self):
        path = self._ask_open(title="Choose guest list", filetypes=EXCEL_TYPES)
        if path and self.load_names(path):
            self.update_ready_status()

    def load_names(self, path, quiet=False):
        try:
            names = cards.read_names(path)
            if not names:
                raise ValueError("Put one name per row in column A of the first sheet.")
        except Exception as e:
            if not quiet:
                messagebox.showerror("Could not read guest list", str(e))
            return False
        self.names, self.names_path = names, str(Path(path).resolve())
        self.names_text.set(f"{Path(path).name}  ·  {plural(len(names), 'guest')}")
        self.preview_menu.configure(values=names)
        self.preview_name.set(names[0])
        self.update_font_hint()
        self.schedule_save()
        return True

    def update_ready_status(self):
        if self.base_image is None:
            self.status.set("Next: choose your invitation image.")
        elif not self.names:
            self.status.set("Next: choose your Excel guest list (or download the template).")
        else:
            self.status.set("Place the name on the preview, then click Generate Cards.")

    def save_template(self):
        path = filedialog.asksaveasfilename(
            title="Save Excel template", defaultextension=".xlsx", initialdir=self.last_dir,
            initialfile="guest_list_template.xlsx", filetypes=EXCEL_TYPES,
        )
        if not path:
            return
        self.last_dir = str(Path(path).parent)
        try:
            cards.write_template(path)
        except Exception as e:
            messagebox.showerror("Could not save template", str(e))
            return
        if messagebox.askyesno("Template saved", f"Saved to {path}.\n\nOpen the folder now?"):
            open_folder(self.last_dir)

    def choose_font(self):
        path = self._ask_open(title="Choose font", filetypes=[("Fonts", "*.ttf *.otf")])
        if not path:
            return
        try:
            ImageFont.truetype(path, 20)
        except Exception as e:
            messagebox.showerror("Could not load font", str(e))
            return
        self.font_label.set(self._add_custom_font(Path(path)).label)

    def _add_custom_font(self, path):
        choice = cards.FontChoice(f"My font: {path.name}", path, chinese=None)
        self.fonts[choice.label] = choice
        self.custom_font_path = path
        self.font_menu.configure(values=list(self.fonts))
        return choice

    @property
    def font(self):
        return self.fonts[self.font_label.get()]

    def update_font_hint(self):
        hint = ""
        if self.font.chinese is False and any(cards.has_chinese(n) for n in self.names):
            fallback = cards.chinese_fallback()
            hint = (f"Names in Chinese will use {fallback.label.split('  ')[0]}." if fallback
                    else "Some names are in Chinese. Choose a font that supports Chinese.")
        self.font_hint.set(hint)

    def choose_color(self):
        _, hex_color = colorchooser.askcolor(color=self.color, title="Text colour")
        if hex_color:
            self._set_color(hex_color)
            self.schedule_preview()
            self.schedule_save()

    def _set_color(self, hex_color):
        self.color = hex_color
        self.color_btn.configure(fg_color=hex_color, hover_color=hex_color)
        self.color_text.set(hex_color.upper())

    def set_placement(self, p):
        for field, var in self.placement_vars.items():
            var.set(str(getattr(p, field)))
        self._set_color(p.color)

    def reset_placement(self):
        self.set_placement(cards.Placement())
        self.after(200, self.scroll_to_name)

    def choose_output(self):
        path = filedialog.askdirectory(
            title="Choose where to save the cards", initialdir=self.output_dir.get() or self.last_dir
        )
        if path:
            self.output_dir.set(path)

    def placement(self):
        try:
            values = {field: int(var.get()) for field, var in self.placement_vars.items()}
        except ValueError:
            return None
        return cards.Placement(**values, color=self.color)

    # --- Preview ---

    def schedule_preview(self):
        if self._preview_job:
            self.after_cancel(self._preview_job)
        self._preview_job = self.after(150, self.render_preview)

    def render_preview(self):
        self._preview_job = None
        if self.base_image is None:
            self.canvas.delete("all")
            self.canvas.configure(scrollregion=(0, 0, 0, 0))
            self.canvas.create_text(
                self.canvas.winfo_width() // 2, self.canvas.winfo_height() // 2, fill=MUTED, font=font(13),
                justify="center", text="Your invitation will appear here.\nStart by choosing an image on the left.",
            )
            return
        p = self.placement()
        if p is None:
            return
        img = cards.render_card(self.base_image, self.preview_name.get() or PREVIEW_PLACEHOLDER, self.font, p)

        # Guides: allowed text area (red) and centre line (blue)
        draw = ImageDraw.Draw(img)
        width = max(2, img.width // 400)
        draw.rectangle([p.x_min, p.base_y, p.x_max, p.max_y + p.max_font_size], outline="#dc2626", width=width)
        draw.line([p.x_center, p.base_y - p.max_font_size // 2, p.x_center, p.max_y + p.max_font_size * 3 // 2],
                  fill="#2563eb", width=width)

        canvas_width = max(self.canvas.winfo_width(), 200)
        self.preview_scale = canvas_width / img.width
        shown = img.resize((canvas_width, round(img.height * self.preview_scale)), Image.LANCZOS)
        self.preview_photo = ImageTk.PhotoImage(shown)
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor="nw", image=self.preview_photo)
        self.canvas.configure(scrollregion=(0, 0, shown.width, shown.height))

    def scroll_to_name(self):
        p = self.placement()
        if self.base_image is None or p is None:
            return
        self.update_idletasks()
        h = self.base_image.height
        self.canvas.yview_moveto(max(0.0, p.base_y / h - self.canvas.winfo_height() / 2 / (h * self.preview_scale)))

    def on_preview_click(self, event):
        p = self.placement()
        if self.base_image is None or p is None:
            return
        x = round(self.canvas.canvasx(event.x) / self.preview_scale)
        y = round(self.canvas.canvasy(event.y) / self.preview_scale)
        half_span = (p.x_max - p.x_min) // 2
        values = {
            "x_center": x,
            "x_min": max(0, x - half_span),
            "x_max": min(self.base_image.width, x + half_span),
            "base_y": y,
            "max_y": y + (p.max_y - p.base_y),
        }
        for field, value in values.items():
            self.placement_vars[field].set(str(value))

    # --- Generate ---

    def generate(self):
        p = self.placement()
        for missing, title, message in [
            (self.base_image is None, "Missing image", "Choose an invitation image first."),
            (not self.names, "Missing guest list", "Choose an Excel guest list first."),
            (p is None, "Check position", "Position values must be whole numbers."),
            (not self.output_dir.get().strip(), "Missing folder", "Choose where to save the cards."),
        ]:
            if missing:
                messagebox.showwarning(title, message)
                return

        out = Path(self.output_dir.get().strip())
        try:
            out.mkdir(parents=True, exist_ok=True)
        except OSError as e:
            messagebox.showerror("Could not create folder", str(e))
            return

        self._jobs, self._total = list(self.names), len(self.names)
        self._out, self._placement, self._font, self._used = out, p, self.font, set()
        self.progress.set(0)
        self.generate_btn.configure(state="disabled", text="Generating…")
        self.save_now()
        # One card per event-loop tick keeps the window responsive
        self.after(10, self._generate_next)

    def _generate_next(self):
        if not self._jobs:
            return self._finish()
        name = self._jobs.pop(0)
        filename = cards.safe_filename(name, self.prefix.get())
        stem, n = filename[:-4], 2
        while filename.lower() in self._used:  # duplicate names get "(2)", "(3)", …
            filename, n = f"{stem} ({n}).jpg", n + 1
        self._used.add(filename.lower())

        self.status.set(f"Creating {len(self._used)} of {self._total}: {name}")
        try:
            card = cards.render_card(self.base_image, name, self._font, self._placement)
            card.save(self._out / filename, quality=95, subsampling=0, optimize=True)
        except Exception as e:
            self._jobs = []
            self.generate_btn.configure(state="normal", text="Generate Cards")
            self.status.set("Stopped because of an error.")
            messagebox.showerror("Could not create card", f"{name}: {e}")
            return
        self.progress.set(len(self._used) / self._total)
        self.after(1, self._generate_next)

    def _finish(self):
        count = len(self._used)
        self.generate_btn.configure(state="normal", text="Generate Cards")
        self.status.set(f"Done! Created {plural(count, 'card')} in {self._out}")
        if messagebox.askyesno("All done", f"Created {plural(count, 'card')}.\n\nOpen the folder now?"):
            open_folder(self._out)


def main():
    ctk.set_appearance_mode("light")
    App().mainloop()


if __name__ == "__main__":
    main()
