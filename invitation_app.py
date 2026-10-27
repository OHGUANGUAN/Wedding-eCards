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
PLACEMENT_FIELDS = [
    ("x_center", "Centre X"),
    ("base_y", "Top Y"),
    ("x_min", "Left limit"),
    ("max_y", "Max Y"),
    ("x_max", "Right limit"),
    ("max_font_size", "Max font size"),
]
PREVIEW_PLACEHOLDER = "Guest Name"
EXCEL_TYPES = [("Excel workbook", "*.xlsx")]

# --- Palette ---
BG = "#F7F1EC"
CARD = "#FFFFFF"
FIELD = "#FBF8F6"
BORDER = "#EADFD6"
ACCENT = "#B5735C"
ACCENT_HOVER = "#9C5E49"
ACCENT_SOFT = "#F3E4DC"
TEXT = "#2F2A27"
MUTED = "#8A7F78"
PREVIEW_BG = "#EDE6E0"


# --- Settings (saved in the user's profile, never in the project folder) ---

def settings_path():
    if sys.platform == "win32":
        base = Path(os.environ.get("APPDATA", Path.home()))
    elif sys.platform == "darwin":
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
    if sys.platform == "win32":
        os.startfile(path)
    elif sys.platform == "darwin":
        subprocess.run(["open", path])
    else:
        subprocess.run(["xdg-open", path])


# --- Styled widgets ---

def font(size=13, weight="normal"):
    return ctk.CTkFont(size=size, weight=weight)


def primary_button(parent, text, command, **kw):
    kw = {"height": 34, "font": font(13, "bold"), **kw}
    return ctk.CTkButton(
        parent, text=text, command=command, fg_color=ACCENT, hover_color=ACCENT_HOVER,
        text_color="white", corner_radius=8, **kw,
    )


def secondary_button(parent, text, command, **kw):
    kw = {"height": 34, "font": font(13), **kw}
    return ctk.CTkButton(
        parent, text=text, command=command, fg_color="transparent", hover_color=ACCENT_SOFT,
        border_width=1, border_color=ACCENT, text_color=ACCENT, corner_radius=8, **kw,
    )


def link_button(parent, text, command):
    return ctk.CTkButton(
        parent, text=text, command=command, fg_color="transparent", hover_color=ACCENT_SOFT,
        text_color=ACCENT, height=26, width=0, anchor="w", font=font(12, "bold"),
    )


def muted_label(parent, **kw):
    kw.setdefault("wraplength", 330)
    return ctk.CTkLabel(parent, text_color=MUTED, font=font(12), anchor="w", justify="left", **kw)


def entry(parent, var, **kw):
    return ctk.CTkEntry(
        parent, textvariable=var, fg_color=FIELD, border_color=BORDER, text_color=TEXT, corner_radius=8, **kw
    )


def option_menu(parent, var, values, **kw):
    return ctk.CTkOptionMenu(
        parent, variable=var, values=values, fg_color=FIELD, button_color=BORDER, button_hover_color=ACCENT_SOFT,
        text_color=TEXT, dropdown_fg_color=CARD, dropdown_hover_color=ACCENT_SOFT, dropdown_text_color=TEXT,
        corner_radius=8, dynamic_resizing=False, anchor="w", font=font(13), **kw,
    )


def step_card(parent, number, title):
    """A white rounded card with a numbered heading; returns the body frame."""
    frame = ctk.CTkFrame(parent, fg_color=CARD, border_color=BORDER, border_width=1, corner_radius=14)
    frame.pack(fill="x", pady=(0, 12), padx=(0, 6))
    header = ctk.CTkFrame(frame, fg_color="transparent")
    header.pack(fill="x", padx=16, pady=(14, 8))
    ctk.CTkLabel(
        header, text=str(number), width=26, height=26, corner_radius=13,
        fg_color=ACCENT, text_color="white", font=font(13, "bold"),
    ).pack(side="left")
    ctk.CTkLabel(header, text=title, text_color=TEXT, font=font(15, "bold")).pack(side="left", padx=(10, 0))
    body = ctk.CTkFrame(frame, fg_color="transparent")
    body.pack(fill="x", padx=16, pady=(0, 14))
    return body


class App(ctk.CTk):
    def __init__(self):
        super().__init__(fg_color=BG)
        self.title(APP_NAME)
        self.minsize(820, 520)

        self.settings = load_settings()
        self.base_image = None
        self.image_path = None
        self.names = []
        self.names_path = None
        self.fonts = {f.label: f for f in cards.available_fonts()}
        self.custom_font_path = None
        self.color = cards.Placement.color
        self.last_dir = self.settings.get("last_dir") or str(Path.home())
        self.preview_photo = None
        self.preview_scale = 1.0
        self._preview_job = None
        self._save_job = None
        self._jobs = []

        defaults = cards.Placement()
        self.placement_vars = {field: tk.StringVar(value=str(getattr(defaults, field))) for field, _ in PLACEMENT_FIELDS}
        self.image_text = tk.StringVar(value="No image chosen yet")
        self.names_text = tk.StringVar(value="No guest list chosen yet")
        self.font_label = tk.StringVar(value=next(iter(self.fonts)))
        self.font_hint = tk.StringVar()
        self.color_text = tk.StringVar(value=self.color.upper())
        self.preview_name = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.prefix = tk.StringVar(value="Invite_")
        self.status = tk.StringVar(value="Choose an invitation image and a guest list to begin.")

        self._build_sidebar()
        self._build_preview()
        self._restore_window()
        self._restore_settings()

        for var in self.placement_vars.values():
            var.trace_add("write", lambda *_: self.schedule_preview())
        self.preview_name.trace_add("write", lambda *_: self.schedule_preview())
        self.font_label.trace_add("write", lambda *_: self.on_font_change())
        for var in [*self.placement_vars.values(), self.font_label, self.output_dir, self.prefix]:
            var.trace_add("write", lambda *_: self.schedule_save())
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    # --- Layout ---

    def _build_sidebar(self):
        sidebar = ctk.CTkFrame(self, fg_color=BG, corner_radius=0)
        sidebar.pack(side="left", fill="y", padx=(16, 8), pady=16)

        header = ctk.CTkFrame(sidebar, fg_color="transparent")
        header.pack(fill="x", pady=(0, 12))
        ctk.CTkLabel(header, text=APP_NAME, text_color=TEXT, font=font(22, "bold")).pack(anchor="w")
        ctk.CTkLabel(
            header, text="A personalised wedding card for every guest", text_color=MUTED, font=font(13)
        ).pack(anchor="w")

        # Packed before the scrolling steps so it always stays visible at the bottom
        footer = ctk.CTkFrame(sidebar, fg_color="transparent")
        footer.pack(side="bottom", fill="x", pady=(12, 0))
        self.generate_btn = primary_button(footer, "Generate Cards", self.generate, height=44, font=font(15, "bold"))
        self.generate_btn.pack(fill="x")
        self.progress = ctk.CTkProgressBar(footer, progress_color=ACCENT, fg_color=BORDER, height=6)
        self.progress.set(0)
        self.progress.pack(fill="x", pady=(10, 6))
        muted_label(footer, textvariable=self.status, wraplength=360).pack(fill="x")

        steps = self.steps = ctk.CTkScrollableFrame(
            sidebar, fg_color=BG, width=370, scrollbar_button_color=BORDER, scrollbar_button_hover_color=ACCENT_SOFT
        )
        steps.pack(fill="both", expand=True)

        body = step_card(steps, 1, "Invitation image")
        primary_button(body, "Choose Image", self.choose_image).pack(anchor="w")
        muted_label(body, textvariable=self.image_text).pack(fill="x", pady=(6, 0))

        body = step_card(steps, 2, "Guest list")
        row = ctk.CTkFrame(body, fg_color="transparent")
        row.pack(fill="x")
        primary_button(row, "Choose Excel File", self.choose_names).pack(side="left")
        secondary_button(row, "Download Excel Template", self.save_template).pack(side="left", padx=(8, 0))
        muted_label(body, textvariable=self.names_text).pack(fill="x", pady=(6, 0))

        body = step_card(steps, 3, "Font")
        self.font_menu = option_menu(body, self.font_label, list(self.fonts))
        self.font_menu.pack(fill="x")
        link_button(body, "Use my own font file…", self.choose_font).pack(anchor="w", pady=(4, 0))
        muted_label(body, textvariable=self.font_hint).pack(fill="x")

        body = step_card(steps, 4, "Name position")
        muted_label(body, text="Click the preview to place the name, or fine-tune in image pixels.").pack(fill="x")
        grid = ctk.CTkFrame(body, fg_color="transparent")
        grid.pack(fill="x", pady=(8, 0))
        for i, (field, label) in enumerate(PLACEMENT_FIELDS):
            r, c = i // 2, (i % 2) * 2
            ctk.CTkLabel(grid, text=label, text_color=TEXT, font=font(12)).grid(row=r, column=c, sticky="w", pady=3)
            entry(grid, self.placement_vars[field], width=72).grid(row=r, column=c + 1, sticky="w", padx=(6, 14), pady=3)
        row = ctk.CTkFrame(body, fg_color="transparent")
        row.pack(fill="x", pady=(8, 0))
        ctk.CTkLabel(row, text="Colour", text_color=TEXT, font=font(12)).pack(side="left")
        self.color_btn = ctk.CTkButton(
            row, text="", width=30, height=26, corner_radius=6, fg_color=self.color, hover_color=self.color,
            border_width=1, border_color=BORDER, command=self.choose_color,
        )
        self.color_btn.pack(side="left", padx=(8, 6))
        ctk.CTkLabel(row, textvariable=self.color_text, text_color=MUTED, font=font(12)).pack(side="left")
        link_button(row, "Reset position", self.reset_placement).pack(side="right")
        ctk.CTkLabel(body, text="Preview with", text_color=TEXT, font=font(12)).pack(anchor="w", pady=(10, 2))
        self.preview_menu = option_menu(body, self.preview_name, [PREVIEW_PLACEHOLDER])
        self.preview_menu.pack(fill="x")

        body = step_card(steps, 5, "Save to")
        row = ctk.CTkFrame(body, fg_color="transparent")
        row.pack(fill="x")
        entry(row, self.output_dir).pack(side="left", fill="x", expand=True)
        secondary_button(row, "Browse", self.choose_output, width=80).pack(side="left", padx=(8, 0))
        row = ctk.CTkFrame(body, fg_color="transparent")
        row.pack(fill="x", pady=(8, 0))
        ctk.CTkLabel(row, text="File name prefix", text_color=TEXT, font=font(12)).pack(side="left")
        entry(row, self.prefix, width=140).pack(side="left", padx=(8, 0))

    def _build_preview(self):
        pane = ctk.CTkFrame(self, fg_color=CARD, border_color=BORDER, border_width=1, corner_radius=14)
        pane.pack(side="left", fill="both", expand=True, padx=(8, 16), pady=16)
        header = ctk.CTkFrame(pane, fg_color="transparent")
        header.pack(fill="x", padx=18, pady=(14, 8))
        ctk.CTkLabel(header, text="Preview", text_color=TEXT, font=font(15, "bold")).pack(side="left")
        ctk.CTkLabel(
            header, text="Red box: space the name must fit in.  Blue line: centre.", text_color=MUTED, font=font(12)
        ).pack(side="right")

        area = ctk.CTkFrame(pane, fg_color="transparent")
        area.pack(fill="both", expand=True, padx=14, pady=(0, 14))
        self.canvas = tk.Canvas(area, background=PREVIEW_BG, highlightthickness=0, cursor="crosshair")
        scroll = ctk.CTkScrollbar(area, command=self.canvas.yview, button_color=BORDER, button_hover_color=ACCENT_SOFT)
        self.canvas.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y", padx=(6, 0))
        self.canvas.pack(side="left", fill="both", expand=True)

        self.canvas.bind("<Configure>", lambda _: self.schedule_preview())
        self.canvas.bind("<Button-1>", self.on_preview_click)
        self.canvas.bind("<MouseWheel>", lambda e: self.canvas.yview_scroll(int(-e.delta / 120), "units"))
        self.canvas.bind("<Button-4>", lambda _: self.canvas.yview_scroll(-1, "units"))
        self.canvas.bind("<Button-5>", lambda _: self.canvas.yview_scroll(1, "units"))

    def _draw_placeholder(self):
        self.canvas.delete("all")
        self.canvas.configure(scrollregion=(0, 0, 0, 0))
        w, h = self.canvas.winfo_width(), self.canvas.winfo_height()
        self.canvas.create_text(
            w // 2, h // 2, fill=MUTED, font=("Segoe UI", 13), justify="center",
            text="Your invitation will appear here.\nStart by choosing an image on the left.",
        )

    # --- Settings ---

    def _restore_window(self):
        self.geometry(self.settings.get("geometry") or "1200x800")
        if self.settings.get("zoomed", sys.platform == "win32"):
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

        reopened = []
        if s.get("image") and Path(s["image"]).exists() and self.load_image(s["image"], quiet=True):
            reopened.append("invitation")
        if s.get("names") and Path(s["names"]).exists() and self.load_names(s["names"], quiet=True):
            reopened.append("guest list")
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
        if not zoomed:
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
            # The last normal-size geometry, so un-maximising restores it
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
        except Exception as e:
            if not quiet:
                messagebox.showerror("Could not read guest list", str(e))
            return False
        if not names:
            if not quiet:
                messagebox.showwarning("No names found", "Put one name per row in column A of the first sheet.")
            return False
        self.names, self.names_path = names, str(Path(path).resolve())
        self.names_text.set(f"{Path(path).name}  ·  {len(names)} guest{'s' if len(names) != 1 else ''}")
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
            open_folder(str(Path(path).parent))

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

    def on_font_change(self):
        self.update_font_hint()
        self.schedule_preview()

    def update_font_hint(self):
        hint = ""
        if self.font.chinese is False and any(cards.has_chinese(n) for n in self.names):
            fallback = cards.chinese_fallback()
            hint = (
                f"Names in Chinese will use {fallback.label.split('  ')[0]}." if fallback
                else "Some names are in Chinese. Choose a font that supports Chinese."
            )
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

    def reset_placement(self):
        defaults = cards.Placement()
        for field, var in self.placement_vars.items():
            var.set(str(getattr(defaults, field)))
        self._set_color(defaults.color)
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
            self._draw_placeholder()
            return
        p = self.placement()
        if p is None:
            return
        img = cards.render_card(self.base_image, self.preview_name.get() or PREVIEW_PLACEHOLDER, self.font, p)

        # Guides: allowed text area (red) and centre line (blue)
        draw = ImageDraw.Draw(img)
        width = max(2, img.width // 400)
        draw.rectangle([p.x_min, p.base_y, p.x_max, p.max_y + p.max_font_size], outline="#dc2626", width=width)
        draw.line(
            [p.x_center, p.base_y - p.max_font_size // 2, p.x_center, p.max_y + p.max_font_size * 3 // 2],
            fill="#2563eb", width=width,
        )

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
        offset = p.base_y / self.base_image.height - self.canvas.winfo_height() / 2 / (self.base_image.height * self.preview_scale)
        self.canvas.yview_moveto(max(0.0, offset))

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
        if self.base_image is None:
            messagebox.showwarning("Missing image", "Choose an invitation image first.")
            return
        if not self.names:
            messagebox.showwarning("Missing guest list", "Choose an Excel guest list first.")
            return
        if p is None:
            messagebox.showwarning("Check position", "Position values must be whole numbers.")
            return
        if not self.output_dir.get().strip():
            messagebox.showwarning("Missing folder", "Choose where to save the cards.")
            return

        out = Path(self.output_dir.get().strip())
        try:
            out.mkdir(parents=True, exist_ok=True)
        except OSError as e:
            messagebox.showerror("Could not create folder", str(e))
            return

        self._jobs = list(self.names)
        self._total = len(self._jobs)
        self._out, self._placement, self._font, self._used = out, p, self.font, set()
        self.progress.set(0)
        self.generate_btn.configure(state="disabled", text="Generating…")
        self.save_now()
        # One card per event-loop tick keeps the window responsive
        self.after(10, self._generate_next)

    def _generate_next(self):
        if not self._jobs:
            self._finish()
            return
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
        self.status.set(f"Done! Created {count} card{'s' if count != 1 else ''} in {self._out}")
        if messagebox.askyesno("All done", f"Created {count} card(s).\n\nOpen the folder now?"):
            open_folder(str(self._out))


def main():
    ctk.set_appearance_mode("light")
    App().mainloop()


if __name__ == "__main__":
    main()
