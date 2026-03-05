#!/usr/bin/env python3
"""PicNRec GUI Tool"""

import datetime
import platform
import threading
from collections.abc import Callable
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from PIL import Image, ImageTk

from picnrec.core import (
    PicNRecDevice, PALETTES, PALETTE_GRAYSCALE,
    decode_gb_camera_image, create_gif, create_mkv,
    IMAGE_WIDTH, IMAGE_HEIGHT, IMAGE_DATA_SIZE,
)


def safe_int(value: object, default: int | None = None) -> int | None:
    if value is None:
        return default
    try:
        s = str(value).strip().replace(',', '')
        if not s:
            return default
        return int(round(float(s)))
    except (ValueError, TypeError):
        return default


class ThemeManager:

    DARK = {
        'bg': '#1e1e1e',
        'bg_secondary': '#252526',
        'bg_tertiary': '#2d2d2d',
        'bg_input': '#3c3c3c',
        'bg_panel': '#282828',
        'fg': '#d4d4d4',
        'fg_secondary': '#9d9d9d',
        'fg_disabled': '#5d5d5d',
        'border': '#505050',
        'border_light': '#606060',
        'border_focus': '#007acc',
        'accent': '#007acc',
        'accent_hover': '#1c97ea',
        'canvas': '#0d0d0d',
        'canvas_border': '#3a3a3a',
        'slider_trough': '#1a1a1a',
    }

    LIGHT = {
        'bg': '#f5f5f5',
        'bg_secondary': '#ffffff',
        'bg_tertiary': '#e8e8e8',
        'bg_input': '#ffffff',
        'bg_panel': '#fafafa',
        'fg': '#1e1e1e',
        'fg_secondary': '#5d5d5d',
        'fg_disabled': '#a0a0a0',
        'border': '#d0d0d0',
        'border_light': '#e0e0e0',
        'border_focus': '#0078d4',
        'accent': '#0078d4',
        'accent_hover': '#106ebe',
        'canvas': '#e8e8e8',
        'canvas_border': '#c0c0c0',
        'slider_trough': '#d0d0d0',
    }

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.current_theme = 'dark'
        self.colors = self.DARK.copy()
        self.style = ttk.Style()
        self.style.theme_use('clam')

    def apply_theme(self, theme_name: str = 'dark') -> dict[str, str]:
        self.current_theme = theme_name
        self.colors = self.DARK.copy() if theme_name == 'dark' else self.LIGHT.copy()
        c = self.colors

        self.style.configure('.',
            background=c['bg'], foreground=c['fg'],
            bordercolor=c['border'], darkcolor=c['bg_tertiary'],
            lightcolor=c['bg_secondary'], troughcolor=c['bg_tertiary'],
            fieldbackground=c['bg_input'], font=('TkDefaultFont', 10))

        self.style.configure('TFrame', background=c['bg'])
        self.style.configure('Toolbar.TFrame', background=c['bg_tertiary'])
        self.style.configure('Sidebar.TFrame', background=c['bg_panel'])

        self.style.configure('TLabelframe',
            background=c['bg_panel'], bordercolor=c['border_light'], relief='solid')
        self.style.configure('TLabelframe.Label',
            background=c['bg_panel'], foreground=c['fg'],
            font=('TkDefaultFont', 9, 'bold'))

        self.style.configure('TLabel', background=c['bg'], foreground=c['fg'])
        self.style.configure('Secondary.TLabel', foreground=c['fg_secondary'])
        self.style.configure('Title.TLabel', font=('TkDefaultFont', 12, 'bold'))
        self.style.configure('Sidebar.TLabel',
            background=c['bg_panel'], foreground=c['fg'])
        self.style.configure('Sidebar.Title.TLabel',
            background=c['bg_panel'], foreground=c['fg'],
            font=('TkDefaultFont', 12, 'bold'))

        self.style.configure('TButton',
            background=c['bg_tertiary'], foreground=c['fg'],
            bordercolor=c['border'], padding=(10, 5))
        self.style.map('TButton',
            background=[('active', c['bg_input']), ('pressed', c['accent'])],
            foreground=[('pressed', '#ffffff')],
            bordercolor=[('focus', c['border_focus'])])

        self.style.configure('Accent.TButton',
            background=c['accent'], foreground='#ffffff')
        self.style.map('Accent.TButton',
            background=[('active', c['accent_hover']), ('pressed', c['accent'])])

        self.style.configure('TEntry',
            fieldbackground=c['bg_input'], foreground=c['fg'],
            bordercolor=c['border'], insertcolor=c['fg'])
        self.style.map('TEntry',
            bordercolor=[('focus', c['border_focus'])],
            lightcolor=[('focus', c['border_focus'])])

        self.style.configure('TCombobox',
            fieldbackground=c['bg_input'], background=c['bg_tertiary'],
            foreground=c['fg'], bordercolor=c['border'], arrowcolor=c['fg'])
        self.style.map('TCombobox',
            fieldbackground=[('readonly', c['bg_input'])],
            bordercolor=[('focus', c['border_focus'])])

        self.style.configure('TSpinbox',
            fieldbackground=c['bg_input'], background=c['bg_tertiary'],
            foreground=c['fg'], bordercolor=c['border'], arrowcolor=c['fg'])

        self.style.configure('TCheckbutton',
            background=c['bg'], foreground=c['fg'])
        self.style.map('TCheckbutton',
            background=[('active', c['bg'])],
            foreground=[('active', c['accent']), ('disabled', c['fg_disabled'])])

        self.style.configure('TRadiobutton',
            background=c['bg'], foreground=c['fg'])
        self.style.map('TRadiobutton',
            background=[('active', c['bg'])],
            foreground=[('active', c['accent']), ('disabled', c['fg_disabled'])])

        self.style.configure('TScale',
            background=c['bg'], troughcolor=c['slider_trough'],
            bordercolor=c['border'])
        self.style.configure('Horizontal.TScale',
            background=c['bg'], troughcolor=c['slider_trough'])

        self.style.configure('TProgressbar',
            background=c['accent'], troughcolor=c['bg_tertiary'],
            bordercolor=c['border'])

        self.style.configure('TSeparator', background=c['border'])

        self.style.configure('TScrollbar',
            background=c['bg_tertiary'], troughcolor=c['bg'],
            bordercolor=c['border'], arrowcolor=c['fg'])

        self.root.option_add('*Menu.background', c['bg_secondary'])
        self.root.option_add('*Menu.foreground', c['fg'])
        self.root.option_add('*Menu.activeBackground', c['accent'])
        self.root.option_add('*Menu.activeForeground', '#ffffff')
        self.root.option_add('*Menu.disabledForeground', c['fg_disabled'])
        self.root.option_add('*Menu.borderWidth', 1)
        self.root.option_add('*Menu.relief', 'flat')
        self.root.option_add('*Toplevel.background', c['bg'])

        return c


def setup_dpi_awareness(root: tk.Tk) -> tuple[float, float]:
    try:
        if platform.system() == 'Windows':
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(2)
        dpi = root.winfo_fpixels('1i')
    except Exception:
        dpi = 96.0

    scale_factor = dpi / 96.0
    if abs(scale_factor - 1.0) > 0.1:
        root.tk.call('tk', 'scaling', scale_factor)

    return scale_factor, dpi


class PalettePreview(ttk.Frame):

    def __init__(
        self,
        parent: tk.Widget,
        palettes: dict,
        on_select: Callable | None = None,
        accent_color: str = '#007acc',
        border_color: str = '#454545',
        **kwargs,
    ) -> None:
        super().__init__(parent, **kwargs)
        self.palettes = palettes
        self.on_select = on_select
        self.accent_color = accent_color
        self.border_color = border_color
        self.selected = tk.StringVar(value='grayscale')
        self.swatch_size = 20
        self.previews: dict[str, tk.Canvas] = {}
        self._create_previews()

    def _create_previews(self) -> None:
        cols = 3
        for i, (name, colors) in enumerate(self.palettes.items()):
            frame = ttk.Frame(self)
            frame.grid(row=i // cols, column=i % cols, padx=4, pady=2)

            canvas = tk.Canvas(
                frame,
                width=self.swatch_size * 4 + 6,
                height=self.swatch_size + 4,
                highlightthickness=2,
                highlightbackground=self.border_color,
            )
            canvas.pack()

            for j, color in enumerate(colors):
                x = j * self.swatch_size + 3
                hex_color = '#{:02x}{:02x}{:02x}'.format(*color)
                canvas.create_rectangle(
                    x, 2, x + self.swatch_size, self.swatch_size + 2,
                    fill=hex_color, outline='')

            label = ttk.Label(frame, text=name.capitalize(),
                              font=('TkDefaultFont', 8))
            label.pack()

            canvas.bind('<Button-1>', lambda e, n=name: self._select(n))
            label.bind('<Button-1>', lambda e, n=name: self._select(n))
            self.previews[name] = canvas

        self._update_selection()

    def _select(self, name: str) -> None:
        self.selected.set(name)
        self._update_selection()
        if self.on_select:
            self.on_select(name)

    def _update_selection(self) -> None:
        selected = self.selected.get()
        for name, canvas in self.previews.items():
            if selected and name == selected:
                canvas.configure(highlightbackground=self.accent_color,
                                 highlightthickness=2)
            else:
                canvas.configure(highlightbackground=self.border_color,
                                 highlightthickness=1)

    def get(self) -> str:
        return self.selected.get()

    def set(self, name: str | None) -> None:
        if name is None:
            self.selected.set('')
        elif name in self.palettes:
            self.selected.set(name)
        self._update_selection()

    def update_colors(self, accent: str, border: str) -> None:
        self.accent_color = accent
        self.border_color = border
        self._update_selection()


class ProgressDialog(tk.Toplevel):

    def __init__(
        self,
        parent: tk.Widget,
        title: str = "Processing",
        message: str = "Please wait...",
        show_stats: bool = False,
    ) -> None:
        super().__init__(parent)
        self.title(title)
        self.transient(parent)
        self.resizable(False, False)
        self.grab_set()

        frame = ttk.Frame(self, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        self.message_var = tk.StringVar(value=message)
        ttk.Label(frame, textvariable=self.message_var,
                  font=('TkDefaultFont', 11, 'bold')).pack(pady=(0, 10))

        self.progress = ttk.Progressbar(frame, length=350, mode='determinate')
        self.progress.pack(pady=(0, 10))

        self.detail_var = tk.StringVar(value="")
        ttk.Label(frame, textvariable=self.detail_var,
                  font=('TkDefaultFont', 10)).pack()

        self.show_stats = show_stats
        if show_stats:
            stats_frame = ttk.Frame(frame)
            stats_frame.pack(pady=(10, 0), fill=tk.X)
            self.stats_var = tk.StringVar(value="")
            ttk.Label(stats_frame, textvariable=self.stats_var,
                      style='Secondary.TLabel',
                      font=('TkMonoFont', 9)).pack()

        self.cancelled = False
        self.cancel_btn = ttk.Button(frame, text="Cancel", command=self._cancel)
        self.cancel_btn.pack(pady=(15, 0))

        self.bind('<Escape>', lambda e: self._cancel())
        self.protocol("WM_DELETE_WINDOW", self._cancel)

        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")

    def _cancel(self) -> None:
        self.cancelled = True
        self.cancel_btn.configure(state='disabled')
        self.message_var.set("Cancelling...")

    def update_progress(
        self,
        value: int,
        maximum: int = 100,
        message: str | None = None,
        detail: str | None = None,
        stats: str | None = None,
    ) -> None:
        self.progress['maximum'] = maximum
        self.progress['value'] = value
        if message is not None:
            self.message_var.set(message)
        if detail is not None:
            self.detail_var.set(detail)
        if stats is not None and self.show_stats:
            self.stats_var.set(stats)
        self.update_idletasks()

    def complete(self) -> None:
        self.grab_release()
        self.destroy()


class ImageNavigator(ttk.Frame):

    def __init__(self, parent: tk.Widget, on_change: Callable | None = None, **kwargs) -> None:
        super().__init__(parent, **kwargs)
        self.on_change = on_change
        self.total = 0
        self.current = 0
        self._create_widgets()

    def _create_widgets(self) -> None:
        self.position_var = tk.StringVar(value="0 / 0")
        ttk.Label(self, textvariable=self.position_var,
                  font=('TkMonoFont', 10), width=12,
                  anchor='center').pack(side=tk.LEFT, padx=(0, 10))

        btn_frame = ttk.Frame(self)
        btn_frame.pack(side=tk.LEFT)

        for text, cmd in [
            ("\u23ee", lambda: self._goto(0)),
            ("\u25c0", lambda: self._step(-1)),
            ("\u25b6", lambda: self._step(1)),
            ("\u23ed", lambda: self._goto(self.total - 1)),
        ]:
            ttk.Button(btn_frame, text=text, width=3,
                       command=cmd).pack(side=tk.LEFT, padx=1)

        self.slider = ttk.Scale(self, from_=0, to=0, orient=tk.HORIZONTAL,
                                command=self._on_slider)
        self.slider.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)

        self.entry_var = tk.StringVar(value="0")
        self.entry = ttk.Entry(self, textvariable=self.entry_var, width=6,
                               justify='center')
        self.entry.pack(side=tk.LEFT)
        self.entry.bind('<Return>', self._on_entry)
        self.entry.bind('<KP_Enter>', self._on_entry)

    def _goto(self, index: int) -> None:
        if self.total == 0:
            return
        index = max(0, min(index, self.total - 1))
        if index != self.current:
            self.current = index
            self._update_display()
            if self.on_change:
                self.on_change(index)

    def _step(self, delta: int) -> None:
        self._goto(self.current + delta)

    def _on_slider(self, value: str) -> None:
        index = int(float(value))
        if index != self.current:
            self.current = index
            self._update_display(update_slider=False)
            if self.on_change:
                self.on_change(index)

    def _on_entry(self, event: tk.Event) -> None:
        index = safe_int(self.entry_var.get())
        if index is not None:
            self._goto(index)

    def _update_display(self, update_slider: bool = True) -> None:
        if self.total > 0:
            self.position_var.set(f"Image {self.current + 1} of {self.total}")
        else:
            self.position_var.set("No images")
        self.entry_var.set(str(self.current))
        if update_slider:
            self.slider.set(self.current)

    def configure_range(self, total: int) -> None:
        self.total = total
        self.current = 0
        self.slider.configure(to=max(0, total - 1))
        self._update_display()

    def get(self) -> int:
        return self.current

    def set(self, index: int) -> None:
        if self.total > 0 and 0 <= index < self.total:
            self.current = index
            self._update_display()


class PicNRecGUI:

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("PicNRec Tk")

        self.scale_factor, self.dpi = setup_dpi_awareness(root)

        base_min_w, base_min_h = 1600, 1000
        base_init_w, base_init_h = 1800, 1100
        self.root.minsize(
            int(base_min_w * self.scale_factor),
            int(base_min_h * self.scale_factor))
        self.root.geometry(
            f"{int(base_init_w * self.scale_factor)}x"
            f"{int(base_init_h * self.scale_factor)}")
        self.root.update_idletasks()

        self.theme = ThemeManager(root)
        self.theme.apply_theme('dark')

        # Device
        self.device: PicNRecDevice | None = None
        self.connected = False

        # Image state
        self.current_image_index = 0
        self.total_images = 0
        self.total_slots = 0
        self.current_image: Image.Image | None = None
        self.image_cache: dict[int, bytes] = {}
        self.valid_indices: list[int] = []
        self.scanned_slots: list[int] = []
        self.failed_slots: list[int] = []
        self.photo_image: ImageTk.PhotoImage | None = None

        # Scan state
        self.dump_thread: threading.Thread | None = None
        self.dump_cancel = False
        self.dump_running = False

        # Export settings
        self.export_range = tk.StringVar(value="1")
        self.export_scale = tk.IntVar(value=1)

        # Settings
        self.palette_name = 'grayscale'
        self.number_padding = tk.IntVar(value=6)
        self.fps = tk.IntVar(value=3)

        # Video metadata
        self.video_title = tk.StringVar(value="")
        self.video_artist = tk.StringVar(value="")
        self.video_comment = tk.StringVar(value="Game Boy Camera timelapse")

        # Custom color adjustment (4-level grayscale)
        self.use_custom_colors = tk.BooleanVar(value=False)
        self.custom_dark = tk.IntVar(value=0)
        self.custom_mid1 = tk.IntVar(value=85)
        self.custom_mid2 = tk.IntVar(value=170)
        self.custom_light = tk.IntVar(value=255)

        # Set automatically when scan used ignore_bitmap (debug mode)
        self.show_all_slots = False

        self._create_menu()
        self._create_toolbar()
        self._create_main_area()

        self.root.configure(bg=self.theme.colors['bg'])
        self._bind_shortcuts()
        self.root.bind('<Configure>', self._on_resize)

    def _bind_shortcuts(self) -> None:
        def is_input_focused() -> bool:
            focused = self.root.focus_get()
            return isinstance(focused, (tk.Entry, ttk.Entry, tk.Text,
                                        ttk.Spinbox, ttk.Scale))

        def nav_left(e: tk.Event) -> None:
            if not is_input_focused():
                self.navigator.set(self.navigator.get() - 1)
                self.load_image(self.navigator.get())

        def nav_right(e: tk.Event) -> None:
            if not is_input_focused():
                self.navigator.set(self.navigator.get() + 1)
                self.load_image(self.navigator.get())

        def nav_home(e: tk.Event) -> None:
            if not is_input_focused():
                self.load_image(0)

        def nav_end(e: tk.Event) -> None:
            if not is_input_focused():
                self.load_image(self.total_images - 1)

        self.root.bind('<Left>', nav_left)
        self.root.bind('<Right>', nav_right)
        self.root.bind('<Home>', nav_home)
        self.root.bind('<End>', nav_end)
        self.root.bind('<Control-o>', lambda e: self.connect_device())
        self.root.bind('<Control-w>', lambda e: self.disconnect_device())
        self.root.bind('<Control-e>', lambda e: self._do_export())
        self.root.bind('<Control-q>', lambda e: self.root.quit())

        def select_all(e: tk.Event) -> str | None:
            w = e.widget
            if isinstance(w, (tk.Entry, ttk.Entry, ttk.Spinbox, tk.Spinbox)):
                w.select_range(0, tk.END)
                w.icursor(tk.END)
                return 'break'
            elif isinstance(w, tk.Text):
                w.tag_add(tk.SEL, '1.0', tk.END)
                w.mark_set(tk.INSERT, tk.END)
                return 'break'
            return None

        for widget_class in ('TEntry', 'Entry', 'TSpinbox', 'Spinbox', 'Text'):
            self.root.bind_class(widget_class, '<Control-a>', select_all)

        def numpad_enter(e: tk.Event) -> str:
            e.widget.event_generate('<Return>')
            return 'break'
        self.root.bind_all('<KP_Enter>', numpad_enter)

    def _create_menu(self) -> None:
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Connect", command=self.connect_device,
                              accelerator="Ctrl+O")
        file_menu.add_command(label="Disconnect", command=self.disconnect_device,
                              accelerator="Ctrl+W")
        file_menu.add_separator()
        file_menu.add_command(label="Export...", command=self._do_export,
                              accelerator="Ctrl+E")
        file_menu.add_separator()
        file_menu.add_command(label="Quit", command=self.root.quit,
                              accelerator="Ctrl+Q")

        nav_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Navigate", menu=nav_menu)
        nav_menu.add_command(label="Previous Image",
                             command=lambda: self._nav_step(-1), accelerator="\u2190")
        nav_menu.add_command(label="Next Image",
                             command=lambda: self._nav_step(1), accelerator="\u2192")
        nav_menu.add_separator()
        nav_menu.add_command(label="First Image",
                             command=lambda: self.load_image(0), accelerator="Home")
        nav_menu.add_command(label="Last Image",
                             command=lambda: self.load_image(self.total_images - 1),
                             accelerator="End")

        device_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Device", menu=device_menu)
        device_menu.add_command(label="Device Info...", command=self.show_device_info)
        device_menu.add_separator()
        device_menu.add_command(label="Re-scan Images", command=self.rescan_images)
        device_menu.add_separator()
        device_menu.add_command(label="Erase All Images...", command=self.erase_images)

        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="View", menu=view_menu)
        self.theme_var = tk.StringVar(value='dark')
        view_menu.add_radiobutton(label="Dark Theme", variable=self.theme_var,
                                  value='dark', command=lambda: self._set_theme('dark'))
        view_menu.add_radiobutton(label="Light Theme", variable=self.theme_var,
                                  value='light', command=lambda: self._set_theme('light'))

        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="User Guide", command=self.show_user_guide)
        help_menu.add_command(label="Keyboard Shortcuts", command=self.show_shortcuts)
        help_menu.add_separator()
        help_menu.add_command(label="About PicNRec Tk", command=self.show_about)

    def _nav_step(self, delta: int) -> None:
        new_idx = self.current_image_index + delta
        if 0 <= new_idx < self.total_images:
            self.load_image(new_idx)

    def _set_theme(self, theme_name: str) -> None:
        self.theme.apply_theme(theme_name)
        self.root.configure(bg=self.theme.colors['bg'])
        self.canvas.configure(bg=self.theme.colors['canvas'])
        self.palette_preview.update_colors(
            self.theme.colors['accent'], self.theme.colors['border'])
        if self.current_image:
            self._display_current_image()

    def _create_toolbar(self) -> None:
        toolbar_container = ttk.Frame(self.root)
        toolbar_container.pack(side=tk.TOP, fill=tk.X)

        tk.Frame(toolbar_container, height=1,
                 bg=self.theme.colors['border']).pack(side=tk.BOTTOM, fill=tk.X)

        toolbar = ttk.Frame(toolbar_container, padding=(10, 8),
                            style='Toolbar.TFrame')
        toolbar.pack(side=tk.TOP, fill=tk.X)

        conn_frame = ttk.Frame(toolbar, style='Toolbar.TFrame')
        conn_frame.pack(side=tk.LEFT)

        self.btn_connect = ttk.Button(conn_frame, text="Connect",
                                      command=self.connect_device,
                                      style='Accent.TButton')
        self.btn_connect.pack(side=tk.LEFT, padx=(0, 5))

        self.btn_scan = ttk.Button(conn_frame, text="Scan",
                                   command=self.start_scan, state='disabled')
        self.btn_scan.pack(side=tk.LEFT, padx=(0, 5))

        self.conn_status = ttk.Label(conn_frame, text="Not connected",
                                     style='Secondary.TLabel')
        self.conn_status.pack(side=tk.LEFT)

        ttk.Label(toolbar,
                  text=f"Scale: {self.scale_factor:.1f}x ({int(self.dpi)} DPI)",
                  style='Secondary.TLabel').pack(side=tk.RIGHT)

    def _create_main_area(self) -> None:
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # Sidebar
        sidebar_container = tk.Frame(main_frame,
                                     bg=self.theme.colors['border'],
                                     padx=1, pady=1)
        sidebar_container.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))

        sidebar = ttk.Frame(sidebar_container, padding=10,
                            style='Sidebar.TFrame')
        sidebar.pack(fill=tk.BOTH, expand=True)

        # Device section
        device_section = ttk.LabelFrame(sidebar, text="Device", padding=10)
        device_section.pack(fill=tk.X, pady=(0, 10))

        self.device_info_label = ttk.Label(device_section, text="Not connected",
                                           style='Secondary.TLabel')
        self.device_info_label.pack(anchor='w')

        self.storage_label = ttk.Label(device_section, text="",
                                       style='Secondary.TLabel')
        self.storage_label.pack(anchor='w')

        # Inline scan progress (hidden by default)
        self.scan_progress_frame = ttk.Frame(device_section)

        self.scan_progress_bar = ttk.Progressbar(self.scan_progress_frame,
                                                  length=200, mode='determinate')
        self.scan_progress_bar.pack(fill=tk.X, pady=(5, 2))

        self.scan_status_label = ttk.Label(self.scan_progress_frame, text="",
                                           font=('TkDefaultFont', 9))
        self.scan_status_label.pack(anchor='w')

        self.scan_detail_label = ttk.Label(self.scan_progress_frame, text="",
                                           style='Secondary.TLabel',
                                           font=('TkMonoFont', 8))
        self.scan_detail_label.pack(anchor='w')

        self.scan_cancel_btn = ttk.Button(self.scan_progress_frame, text="Cancel",
                                          command=self._cancel_scan)
        self.scan_cancel_btn.pack(anchor='w', pady=(5, 0))

        # Palette
        ttk.Label(sidebar, text="Palette",
                  style='Sidebar.Title.TLabel').pack(anchor='w', pady=(5, 5))

        self.palette_preview = PalettePreview(
            sidebar, PALETTES,
            on_select=self._on_palette_change,
            accent_color=self.theme.colors['accent'],
            border_color=self.theme.colors['border'],
        )
        self.palette_preview.pack(pady=(0, 10), anchor='w')

        # Color adjustment
        color_section = ttk.LabelFrame(sidebar, text="Color Adjustment", padding=10)
        color_section.pack(fill=tk.X, pady=(0, 10))
        self._create_color_sliders(color_section)

        # Export section
        export_section = ttk.LabelFrame(sidebar, text="Export", padding=10)
        export_section.pack(fill=tk.X, pady=(0, 10))

        fmt_frame = ttk.Frame(export_section)
        fmt_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(fmt_frame, text="Format:").pack(side=tk.LEFT)
        self.export_format = tk.StringVar(value='bmp')
        ttk.Combobox(fmt_frame, textvariable=self.export_format,
                     values=['bmp', 'gif', 'mkv'], width=6,
                     state='readonly').pack(side=tk.LEFT, padx=(5, 0))

        range_frame = ttk.Frame(export_section)
        range_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(range_frame, text="Range:").pack(side=tk.LEFT)
        ttk.Entry(range_frame, textvariable=self.export_range,
                  width=12).pack(side=tk.LEFT, padx=(5, 0))
        ttk.Label(range_frame, text="(1-50,100)", style='Secondary.TLabel',
                  font=('TkDefaultFont', 8)).pack(side=tk.LEFT, padx=(3, 0))

        scale_frame = ttk.Frame(export_section)
        scale_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(scale_frame, text="Scale:").pack(side=tk.LEFT)
        ttk.Spinbox(scale_frame, from_=1, to=16,
                     textvariable=self.export_scale, width=4
                     ).pack(side=tk.LEFT, padx=(5, 0))
        ttk.Label(scale_frame, text="x", style='Secondary.TLabel').pack(side=tk.LEFT)

        fps_frame = ttk.Frame(export_section)
        fps_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(fps_frame, text="FPS:").pack(side=tk.LEFT)
        ttk.Spinbox(fps_frame, from_=1, to=60,
                     textvariable=self.fps, width=4
                     ).pack(side=tk.LEFT, padx=(5, 0))
        ttk.Label(fps_frame, text="(video)", style='Secondary.TLabel',
                  font=('TkDefaultFont', 8)).pack(side=tk.LEFT, padx=(3, 0))

        ttk.Label(export_section, text="MKV Metadata:",
                  style='Secondary.TLabel').pack(anchor='w', pady=(5, 2))

        meta_grid = ttk.Frame(export_section)
        meta_grid.pack(fill=tk.X)
        for row, (label, var) in enumerate([
            ("Title:", self.video_title),
            ("Artist:", self.video_artist),
            ("Comment:", self.video_comment),
        ]):
            ttk.Label(meta_grid, text=label,
                      font=('TkDefaultFont', 8)).grid(row=row, column=0, sticky='w')
            ttk.Entry(meta_grid, textvariable=var,
                      width=15).grid(row=row, column=1, sticky='ew', padx=(3, 0))
        meta_grid.columnconfigure(1, weight=1)

        ttk.Button(export_section, text="Export...",
                   command=self._do_export).pack(fill=tk.X, pady=(10, 0))

        ttk.Button(sidebar, text="Erase All Images...",
                   command=self.erase_images).pack(fill=tk.X, pady=(10, 0))

        # Image display area
        image_container = tk.Frame(main_frame,
                                   bg=self.theme.colors['canvas_border'],
                                   padx=2, pady=2)
        image_container.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        image_frame = ttk.Frame(image_container)
        image_frame.pack(fill=tk.BOTH, expand=True)

        self.status_banner = tk.Label(
            image_frame, text="Ready",
            bg=self.theme.colors['canvas'], fg='#ffffff',
            font=('TkDefaultFont', 14, 'bold'), pady=8)
        self.status_banner.pack(fill=tk.X)

        self.canvas = tk.Canvas(image_frame, bg=self.theme.colors['canvas'],
                                highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self.canvas.create_text(
            400, 250,
            text="Connect to device to view images\n\nCtrl+O to connect",
            fill='#666666', font=('TkDefaultFont', 12),
            justify='center', tags='placeholder')

        nav_frame = ttk.Frame(image_frame, padding=(0, 10, 0, 0))
        nav_frame.pack(fill=tk.X)
        self.navigator = ImageNavigator(nav_frame, on_change=self.load_image)
        self.navigator.pack(fill=tk.X)

    def _on_resize(self, event: tk.Event) -> None:
        if self.current_image and event.widget == self.root:
            if hasattr(self, '_resize_after_id'):
                self.root.after_cancel(self._resize_after_id)
            self._resize_after_id = self.root.after(
                50, self._display_current_image)

    def _on_palette_change(self, name: str) -> None:
        self.palette_name = name
        self.use_custom_colors.set(False)
        self.refresh_image()

    def _create_color_sliders(self, parent: ttk.Frame) -> None:
        ttk.Checkbutton(parent, text="Use custom colors",
                        variable=self.use_custom_colors,
                        command=self._on_custom_colors_toggle
                        ).pack(anchor='w', pady=(0, 10))

        sliders_frame = ttk.Frame(parent)
        sliders_frame.pack(fill=tk.X)

        slider_configs = [
            ("Light:", self.custom_light),
            ("Mid-1:", self.custom_mid2),
            ("Mid-2:", self.custom_mid1),
            ("Dark:", self.custom_dark),
        ]

        self.color_sliders: list[tuple[ttk.Scale, tk.Canvas, tk.IntVar, ttk.Label]] = []
        for label_text, var in slider_configs:
            row = ttk.Frame(sliders_frame)
            row.pack(fill=tk.X, pady=2)

            ttk.Label(row, text=label_text, width=6).pack(side=tk.LEFT)

            swatch = tk.Canvas(row, width=20, height=20,
                               highlightthickness=1,
                               highlightbackground=self.theme.colors['border'])
            swatch.pack(side=tk.LEFT, padx=(0, 5))

            slider = ttk.Scale(
                row, from_=0, to=255, orient=tk.HORIZONTAL, variable=var,
                command=lambda v, s=swatch, vr=var: self._on_color_slider_change(s, vr))
            slider.pack(side=tk.LEFT, fill=tk.X, expand=True)

            def slider_up(e: tk.Event, v: tk.IntVar = var) -> str:
                v.set(min(255, safe_int(v.get(), 0) + 1))
                return 'break'

            def slider_down(e: tk.Event, v: tk.IntVar = var) -> str:
                v.set(max(0, safe_int(v.get(), 0) - 1))
                return 'break'

            slider.bind('<Up>', slider_up)
            slider.bind('<Down>', slider_down)

            value_label = ttk.Label(row, width=4)
            value_label.pack(side=tk.LEFT, padx=(5, 0))

            var.trace_add('write', lambda *args, s=swatch, vr=var, lbl=value_label:
                          self._update_slider_display(s, vr, lbl))

            self._update_swatch(swatch, var)
            self.color_sliders.append((slider, swatch, var, value_label))

        for _, swatch, var, value_label in self.color_sliders:
            self._update_slider_display(swatch, var, value_label)

    def _update_swatch(self, swatch: tk.Canvas, var: tk.IntVar) -> None:
        try:
            val = int(float(var.get()))
            swatch.configure(bg=f'#{val:02x}{val:02x}{val:02x}')
        except (ValueError, tk.TclError):
            pass

    def _on_color_slider_change(self, swatch: tk.Canvas, var: tk.IntVar) -> None:
        self._update_swatch(swatch, var)
        if self.use_custom_colors.get():
            self.refresh_image()

    def _update_slider_display(
        self, swatch: tk.Canvas, var: tk.IntVar, label: ttk.Label,
    ) -> None:
        try:
            val = int(float(var.get()))
            swatch.configure(bg=f'#{val:02x}{val:02x}{val:02x}')
            label.configure(text=str(val))
        except (ValueError, tk.TclError):
            pass

    def _on_custom_colors_toggle(self) -> None:
        if self.use_custom_colors.get():
            self.palette_preview.set(None)
        self.refresh_image()

    def set_status(self, message: str) -> None:
        self.status_banner.config(text=message)
        self.root.update_idletasks()

    # connection
    def connect_device(self) -> None:
        if self.connected:
            return

        self.set_status("Connecting...")
        self.root.update_idletasks()
        self.device = PicNRecDevice()

        try:
            self.device.connect()
            self.connected = True
            self.btn_connect.config(text="Disconnect", command=self.disconnect_device)
            self.btn_scan.config(state='normal')
            self.conn_status.config(text=f"Connected: {self.device.port}")
            self.device_info_label.config(
                text=f"Detected: {self.device.port}\nSettings: ASCII, 1M baud")
            self.storage_label.config(text="")
            self.set_status("Connected. Click Scan to load images.")
        except Exception as e:
            messagebox.showerror("Connection Error", str(e))
            self.set_status("Connection failed")
            self.device = None

    def disconnect_device(self) -> None:
        if self.dump_running:
            self.dump_cancel = True
            if self.dump_thread and self.dump_thread.is_alive():
                self.dump_thread.join(timeout=1.0)
            self.dump_running = False

        if self.device:
            self.device.disconnect()
        self.device = None
        self.connected = False
        self.btn_connect.config(text="Connect", command=self.connect_device)
        self.btn_scan.config(state='disabled')
        self.conn_status.config(text="Not connected")
        self.device_info_label.config(text="Not connected")
        self.storage_label.config(text="")
        self.total_images = 0
        self.total_slots = 0
        self.current_image_index = 0
        self.current_image = None
        self.image_cache = {}
        self.valid_indices = []
        self.scanned_slots = []
        self.failed_slots = []
        self.show_all_slots = False
        self.navigator.configure_range(0)
        self.canvas.delete('image')
        self.canvas.delete('placeholder')
        self.canvas.create_text(
            self.canvas.winfo_width() // 2,
            self.canvas.winfo_height() // 2,
            text="Connect to device to view images\n\nCtrl+O to connect",
            fill='#666666', font=('TkDefaultFont', 12),
            justify='center', tags='placeholder')
        self.set_status("Disconnected")

    # scanning
    def start_scan(self) -> None:
        if not self.connected:
            return
        if self.dump_running:
            messagebox.showwarning("Scan Running", "Cancel current scan first.")
            return

        self.set_status("Reading bitmap...")
        self.root.update_idletasks()

        try:
            bitmap = self.device.read_bitmap()
            if not bitmap:
                messagebox.showerror("Bitmap Error",
                    "Failed to read device bitmap.\n\n"
                    "Check USB connection (use USB-A to USB-C cable).")
                self.set_status("Bitmap read failed")
                return

            filled_slots = self.device.get_filled_slots(bitmap)
            self._show_scan_dialog(filled_slots)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read bitmap: {e}")
            self.set_status("Bitmap read failed")

    def _resolve_scan_slots(
        self,
        choice: str,
        filled_slots: list[int],
        range_start: int,
        range_end: int,
        adv_spec: str,
        placeholder: str,
        ignore_bitmap: bool,
    ) -> tuple[list[int] | None, str | None]:
        if choice == 'all':
            if ignore_bitmap:
                max_slot = max(filled_slots) if filled_slots else 255
                return list(range(max_slot + 1)), None
            elif filled_slots:
                return filled_slots, None
            else:
                return None, "No images to extract."

        elif choice == 'range':
            start = safe_int(range_start, 1) - 1
            end = safe_int(range_end, 1) - 1
            if start < 0 or end < start:
                return None, "Invalid range. Start must be <= end."
            if ignore_bitmap:
                return list(range(start, end + 1)), None
            slots = [s for s in filled_slots if start <= s <= end]
            if slots:
                return slots, None
            return None, f"No filled slots in range {start + 1}-{end + 1}."

        elif choice == 'advanced':
            if adv_spec == placeholder or not adv_spec.strip():
                return None, "Please enter a range specification."
            requested, parse_error = self._parse_slot_spec(adv_spec)
            if parse_error:
                return None, parse_error

            if not ignore_bitmap:
                max_filled = max(filled_slots) if filled_slots else 0
                reasonable_max = max(max_filled + 100, 500)
                out_of_bounds = [s + 1 for s in requested if s > reasonable_max]
                if out_of_bounds:
                    examples = ", ".join(str(x) for x in out_of_bounds[:3])
                    if len(out_of_bounds) > 3:
                        examples += f", ... (+{len(out_of_bounds) - 3} more)"
                    return None, (
                        f"Slots out of range: {examples}. "
                        f"Highest filled: {max_filled + 1}. "
                        "Typo? Use debug mode to override.")

            if ignore_bitmap:
                return requested, None
            slots = [s for s in filled_slots if s in set(requested)]
            if slots:
                return slots, None
            return None, "No filled slots match your specification."

        return None, "Unknown option."

    def _show_scan_dialog(self, filled_slots: list[int]) -> None:
        dialog = tk.Toplevel(self.root)
        dialog.title("Extract Images")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.configure(bg=self.theme.colors['bg'])

        frame = ttk.Frame(dialog, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="Device Bitmap Status",
                  style='Title.TLabel').pack(pady=(0, 15))

        total_filled = len(filled_slots)
        if total_filled == 0:
            summary = "No images found on device."
            ranges_text = "Device is empty."
        else:
            ranges = self._find_ranges(filled_slots)
            ranges_text = self._format_ranges(ranges)
            summary = f"{total_filled} images found"

        ttk.Label(frame, text=summary, style='Secondary.TLabel',
                  font=('TkDefaultFont', 11, 'bold')).pack()

        # Bitmap visualization
        if total_filled > 0:
            viz_frame = ttk.Frame(frame)
            viz_frame.pack(pady=10, fill=tk.X)

            canvas = tk.Canvas(viz_frame, width=400, height=30,
                               bg=self.theme.colors['bg_input'],
                               highlightthickness=1,
                               highlightbackground=self.theme.colors['border'])
            canvas.pack()

            max_slot = max(filled_slots)
            scale = 400 / max(max_slot + 1, 1)
            for slot in filled_slots:
                x = int(slot * scale)
                canvas.create_line(x, 0, x, 30, fill=self.theme.colors['accent'])

            ttk.Label(viz_frame, text=f"Slots 1 to {max_slot + 1}",
                      style='Secondary.TLabel').pack()

        ttk.Label(frame, text=ranges_text, style='Secondary.TLabel',
                  justify=tk.LEFT).pack(pady=(5, 15))

        ttk.Separator(frame, orient='horizontal').pack(fill=tk.X, pady=10)

        option_var = tk.StringVar(value='all')

        opt_all = ttk.Radiobutton(frame, text=f"Extract all {total_filled} images",
                                  variable=option_var, value='all')
        opt_all.pack(anchor=tk.W, pady=2)
        if total_filled == 0:
            opt_all.configure(state='disabled')

        # Range option (1-indexed for user)
        range_frame = ttk.Frame(frame)
        range_frame.pack(anchor=tk.W, pady=2, fill=tk.X)

        ttk.Radiobutton(range_frame, text="Extract range: slots",
                        variable=option_var, value='range').pack(side=tk.LEFT)

        range_start = tk.IntVar(value=(min(filled_slots) + 1) if filled_slots else 1)
        range_end = tk.IntVar(value=(max(filled_slots) + 1) if filled_slots else 1)

        start_spin = ttk.Spinbox(range_frame, from_=1, to=18720,
                                 textvariable=range_start, width=6)
        start_spin.pack(side=tk.LEFT, padx=2)
        ttk.Label(range_frame, text="to").pack(side=tk.LEFT)
        end_spin = ttk.Spinbox(range_frame, from_=1, to=18720,
                               textvariable=range_end, width=6)
        end_spin.pack(side=tk.LEFT, padx=2)
        ttk.Label(range_frame, text="(only filled slots in range)"
                  ).pack(side=tk.LEFT, padx=5)

        start_spin.bind('<FocusIn>', lambda e: option_var.set('range'))
        end_spin.bind('<FocusIn>', lambda e: option_var.set('range'))

        # Advanced (multi-range)
        adv_frame = ttk.Frame(frame)
        adv_frame.pack(anchor=tk.W, pady=2, fill=tk.X)

        ttk.Radiobutton(adv_frame, text="Multiple ranges:",
                        variable=option_var, value='advanced').pack(side=tk.LEFT)

        adv_entry = ttk.Entry(adv_frame, width=30)
        adv_entry.pack(side=tk.LEFT, padx=5)

        placeholder = "1-50,100-200,250"
        adv_entry.insert(0, placeholder)
        adv_entry.config(foreground='gray')

        def on_focus_in(e: tk.Event) -> None:
            option_var.set('advanced')
            if adv_entry.get() == placeholder:
                adv_entry.delete(0, tk.END)
                adv_entry.config(foreground=self.theme.colors['fg'])

        def on_focus_out(e: tk.Event) -> None:
            if not adv_entry.get():
                adv_entry.insert(0, placeholder)
                adv_entry.config(foreground='gray')

        adv_entry.bind('<FocusIn>', on_focus_in)
        adv_entry.bind('<FocusOut>', on_focus_out)

        ttk.Separator(frame, orient='horizontal').pack(fill=tk.X, pady=10)
        debug_mode = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame,
            text="Debug: ignore bitmap (try to read specified slots directly)",
            variable=debug_mode).pack(anchor=tk.W)

        # Error display
        error_var = tk.StringVar(value="")
        ttk.Label(frame, textvariable=error_var,
                  foreground='red', wraplength=350).pack(pady=(5, 0))

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(pady=(20, 0))

        def do_extract() -> None:
            try:
                sv = range_start.get()
                ev = range_end.get()
            except tk.TclError:
                sv, ev = 1, 1

            slots, error = self._resolve_scan_slots(
                choice=option_var.get(),
                filled_slots=filled_slots,
                range_start=sv,
                range_end=ev,
                adv_spec=adv_entry.get(),
                placeholder=placeholder,
                ignore_bitmap=debug_mode.get(),
            )
            if error:
                error_var.set(error)
                return
            self.show_all_slots = debug_mode.get()
            dialog.destroy()
            self._start_scan_with_slots(slots)

        ttk.Button(btn_frame, text="Extract", command=do_extract,
                   style='Accent.TButton').pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel",
                   command=dialog.destroy).pack(side=tk.LEFT, padx=5)

        dialog.bind('<Return>', lambda e: do_extract())
        dialog.bind('<KP_Enter>', lambda e: do_extract())
        dialog.bind('<Escape>', lambda e: dialog.destroy())

        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - dialog.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{x}+{y}")

    def _find_ranges(self, slots: list[int]) -> list[tuple[int, int]]:
        if not slots:
            return []
        sorted_slots = sorted(slots)
        ranges = []
        start = end = sorted_slots[0]
        for slot in sorted_slots[1:]:
            if slot == end + 1:
                end = slot
            else:
                ranges.append((start, end))
                start = end = slot
        ranges.append((start, end))
        return ranges

    def _format_ranges(self, ranges: list[tuple[int, int]]) -> str:
        if not ranges:
            return "No filled slots"
        if len(ranges) == 1:
            start, end = ranges[0]
            if start == end:
                return f"Slot {start + 1}"
            return f"Slots {start + 1}-{end + 1}"
        parts = []
        for start, end in ranges[:5]:
            if start == end:
                parts.append(str(start + 1))
            else:
                parts.append(f"{start + 1}-{end + 1}")
        result = "Ranges: " + ", ".join(parts)
        if len(ranges) > 5:
            result += f" ... (+{len(ranges) - 5} more)"
        return result

    def _parse_slot_spec(self, spec: str) -> tuple[list[int] | None, str | None]:
        # '1-100,150,200-250' (1-indexed input) -> 0-indexed slots
        slots: list[int] = []
        for part in spec.split(','):
            part = part.strip()
            if not part:
                continue
            if '-' in part and not part.startswith('-') and part.count('-') == 1:
                pieces = part.split('-')
                start_idx = safe_int(pieces[0])
                end_idx = safe_int(pieces[1])
                if start_idx is None or end_idx is None:
                    return None, f"Invalid number in '{part}'"
                start_idx -= 1
                end_idx -= 1
                if start_idx < 0:
                    return None, "Slot numbers start at 1, not 0"
                if end_idx < start_idx:
                    return None, f"Range '{part}' is backwards (end < start)"
                if end_idx - start_idx + 1 > 20000:
                    return None, f"Range {part} is too large (max 18720 slots)"
                slots.extend(range(start_idx, end_idx + 1))
            else:
                val = safe_int(part)
                if val is None:
                    return None, f"Invalid number '{part}'"
                val -= 1
                if val < 0:
                    return None, "Slot numbers start at 1, not 0"
                slots.append(val)
        if not slots:
            return None, "No slots specified"
        return sorted(set(slots)), None

    def _cancel_scan(self) -> None:
        self.dump_cancel = True
        self.scan_cancel_btn.config(state='disabled')
        self.scan_status_label.config(text="Cancelling...")

    def _show_scan_progress(self, show: bool = True) -> None:
        if show:
            self.scan_progress_frame.pack(fill=tk.X, pady=(5, 0))
        else:
            self.scan_progress_frame.pack_forget()

    def _update_scan_progress(
        self, current: int, total: int, slot: int,
        attempt: int, max_attempts: int, valid: int, failed: int,
    ) -> None:
        self.scan_progress_bar['maximum'] = total
        self.scan_progress_bar['value'] = current
        attempt_str = f" (attempt {attempt}/{max_attempts})" if attempt > 1 else ""
        self.scan_status_label.config(text=f"Slot {slot + 1}{attempt_str}")
        self.scan_detail_label.config(
            text=f"{current}/{total}  |  Valid: {valid}  Failed: {failed}")

    def _start_scan_with_slots(self, slots: list[int]) -> None:
        if not self.connected or self.dump_running:
            return

        self.dump_cancel = False
        self.dump_running = True

        self.scan_cancel_btn.config(state='normal')
        self._show_scan_progress(True)
        self.scan_status_label.config(text=f"Starting scan of {len(slots)} slots...")
        self.scan_detail_label.config(text="")

        def scan_thread() -> None:
            self._scan_slots(slots)
            self.root.after(0, self._on_scan_complete)

        self.dump_thread = threading.Thread(target=scan_thread, daemon=True)
        self.dump_thread.start()

    def _scan_slots(self, slots: list[int]) -> None:  # runs in background thread
        self.image_cache = {}
        self.valid_indices = []
        self.scanned_slots = []
        self.failed_slots = []

        for scanned, slot in enumerate(slots, 1):
            if self.dump_cancel:
                break

            self.scanned_slots.append(slot)

            success = False
            max_attempts = 3
            for attempt in range(1, max_attempts + 1):
                if self.dump_cancel:
                    break

                self.root.after(
                    0, lambda s=scanned, t=len(slots), sl=slot,
                    a=attempt, ma=max_attempts,
                    v=len(self.valid_indices), f=len(self.failed_slots):
                    self._update_scan_progress(s, t, sl, a, ma, v, f))

                self.device.soft_reconnect()

                try:
                    data = self.device.read_image_data(slot)
                    if data and len(data) >= IMAGE_DATA_SIZE:
                        self.image_cache[slot] = data
                        self.valid_indices.append(slot)
                        success = True
                        break
                    elif data:
                        self.image_cache[slot] = data
                except Exception:
                    pass

            if not success and not self.dump_cancel:
                self.failed_slots.append(slot)
                print(f"[PicNRec] Slot {slot + 1} failed after {max_attempts} attempts",
                      flush=True)

        self.total_slots = len(self.scanned_slots)
        self.total_images = (
            self.total_slots if self.show_all_slots
            else len(self.valid_indices))

    def _on_scan_complete(self) -> None:
        self.dump_running = False
        self._show_scan_progress(False)

        if self.show_all_slots:
            self.navigator.configure_range(self.total_slots)
        else:
            self.navigator.configure_range(len(self.valid_indices))

        if self.valid_indices:
            self.current_image_index = 0
            first_slot = (self.scanned_slots[0] if self.show_all_slots
                          else self.valid_indices[0])
            if first_slot in self.image_cache:
                self._render_cached_image(first_slot)
            self.export_range.set(f"1-{len(self.valid_indices)}")
        else:
            self.canvas.delete('placeholder')
            self.canvas.create_text(
                self.canvas.winfo_width() // 2,
                self.canvas.winfo_height() // 2,
                text="No valid images found",
                fill='#666666', font=('TkDefaultFont', 12), tags='placeholder')

        cancelled = " (cancelled)" if self.dump_cancel else ""
        failed_count = len(self.failed_slots)

        if failed_count > 0:
            self.set_status(
                f"Found {len(self.valid_indices)} images, "
                f"{failed_count} failed{cancelled}")
            self.storage_label.config(
                text=f"Images: {len(self.valid_indices)} ({failed_count} failed)")
            failed_display = [s + 1 for s in self.failed_slots[:10]]
            suffix = (f" ... +{len(self.failed_slots) - 10} more"
                      if len(self.failed_slots) > 10 else "")
            print(f"[PicNRec] Failed slots (1-indexed): {failed_display}{suffix}")
        else:
            self.set_status(f"Found {len(self.valid_indices)} images{cancelled}")
            self.storage_label.config(text=f"Images: {len(self.valid_indices)}")

    # display
    def _get_slot_for_position(self, position: int) -> int:
        if self.show_all_slots:
            if 0 <= position < len(self.scanned_slots):
                return self.scanned_slots[position]
        else:
            if 0 <= position < len(self.valid_indices):
                return self.valid_indices[position]
        return -1

    def load_image(self, position: int) -> None:
        if position < 0 or position >= self.total_images:
            return
        slot = self._get_slot_for_position(position)
        if slot < 0:
            return
        self.current_image_index = position
        self.navigator.set(position)
        self._load_slot(slot, position)

    def _load_slot(self, slot: int, display_position: int) -> None:
        if slot not in self.image_cache:
            self.set_status(f"Slot {slot + 1} (no data)")
            self.canvas.delete('image')
            self.canvas.create_text(
                self.canvas.winfo_width() // 2,
                self.canvas.winfo_height() // 2,
                text=f"Slot {slot + 1}\n(no data)",
                fill='#666666', font=('TkDefaultFont', 12),
                justify='center', tags='image')
            return

        self._render_cached_image(slot)
        data_len = len(self.image_cache[slot])
        status = "valid" if data_len >= IMAGE_DATA_SIZE else f"partial ({data_len} bytes)"
        self.set_status(f"Image {display_position + 1} (slot {slot + 1}): {status}")

    def _get_current_palette(self) -> list[tuple[int, int, int]]:
        if self.use_custom_colors.get():
            return [
                (max(0, min(255, safe_int(self.custom_light.get(), 255))),) * 3,
                (max(0, min(255, safe_int(self.custom_mid2.get(), 170))),) * 3,
                (max(0, min(255, safe_int(self.custom_mid1.get(), 85))),) * 3,
                (max(0, min(255, safe_int(self.custom_dark.get(), 0))),) * 3,
            ]
        return PALETTES.get(self.palette_name, PALETTE_GRAYSCALE)

    def _render_cached_image(self, slot: int) -> None:
        if slot not in self.image_cache:
            return
        self.current_image = decode_gb_camera_image(
            self.image_cache[slot], self._get_current_palette())
        self._display_current_image()

    def _display_current_image(self) -> None:
        if not self.current_image:
            return

        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()
        if canvas_w < 10 or canvas_h < 10:
            return

        # Integer scaling for pixel art
        scale = max(1, min(canvas_w // IMAGE_WIDTH, canvas_h // IMAGE_HEIGHT))
        display_image = self.current_image.resize(
            (IMAGE_WIDTH * scale, IMAGE_HEIGHT * scale), Image.NEAREST)

        self.photo_image = ImageTk.PhotoImage(display_image)
        self.canvas.delete('image')
        self.canvas.delete('placeholder')
        self.canvas.create_image(
            canvas_w // 2, canvas_h // 2,
            image=self.photo_image, anchor=tk.CENTER, tags='image')

    def refresh_image(self) -> None:
        slot = self._get_slot_for_position(self.current_image_index)
        if slot >= 0 and slot in self.image_cache:
            self._render_cached_image(slot)

    # export
    def _parse_export_range(self, spec: str) -> list[int] | None:
        # '1-5, 10-20' (1-indexed gallery positions) -> 0-indexed
        positions: list[int] = []
        for part in spec.split(','):
            part = part.strip()
            if not part:
                continue
            if '-' in part and not part.startswith('-'):
                pieces = part.split('-', 1)
                start = safe_int(pieces[0])
                end = safe_int(pieces[1])
                if start is None or end is None:
                    return None
                start -= 1
                end -= 1
                if start >= 0 and end >= start:
                    positions.extend(range(start, end + 1))
                else:
                    return None
            else:
                pos = safe_int(part)
                if pos is None:
                    return None
                pos -= 1
                if pos >= 0:
                    positions.append(pos)
        return sorted(set(positions)) if positions else None

    def _do_export(self) -> None:
        if not self.connected or not self.valid_indices:
            messagebox.showwarning("No Images", "No valid images to export.")
            return

        spec = self.export_range.get().strip()
        if not spec:
            indices = self.valid_indices[:]
        else:
            positions = self._parse_export_range(spec)
            if positions is None:
                messagebox.showerror("Invalid Range",
                    "Could not parse range.\nFormat: 1-50,100-200")
                return
            max_pos = len(self.valid_indices)
            out_of_bounds = [p + 1 for p in positions if p >= max_pos]
            if out_of_bounds:
                examples = ", ".join(str(x) for x in out_of_bounds[:3])
                if len(out_of_bounds) > 3:
                    examples += f", ... (+{len(out_of_bounds) - 3} more)"
                messagebox.showerror("Invalid Range",
                    f"Positions out of range: {examples}\n\n"
                    f"Valid range is 1-{max_pos} ({max_pos} images loaded)")
                return
            indices = [self.valid_indices[p] for p in positions if p < max_pos]

        fmt = self.export_format.get()

        if fmt == 'bmp':
            output_dir = filedialog.askdirectory(title="Select Export Directory")
            if not output_dir:
                return
            timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            export_dir = Path(output_dir) / f"picnrec_{timestamp}"
            self._run_export(indices, export_dir)

        elif fmt == 'gif':
            output_path = filedialog.asksaveasfilename(
                title="Save GIF As", defaultextension=".gif",
                filetypes=[("GIF files", "*.gif")])
            if not output_path:
                return
            export_dir = Path(output_path).with_suffix('') / '_frames'
            self._run_export(indices, export_dir, gif_path=Path(output_path))

        elif fmt == 'mkv':
            output_path = filedialog.asksaveasfilename(
                title="Save MKV As", defaultextension=".mkv",
                filetypes=[("MKV files", "*.mkv"), ("All files", "*.*")])
            if not output_path:
                return
            export_dir = Path(output_path).with_suffix('') / '_frames'
            self._run_export(indices, export_dir, mkv_path=Path(output_path))

    def _run_export(
        self,
        indices: list[int],
        output_dir: Path,
        gif_path: Path | None = None,
        mkv_path: Path | None = None,
    ) -> None:
        total = len(indices)
        progress = ProgressDialog(self.root, "Exporting",
                                  f"Exporting {total} images...")

        # Capture tkinter vars on main thread before spawning worker
        palette = self._get_current_palette()
        scale = max(1, safe_int(self.export_scale.get(), 1))
        padding = max(1, safe_int(self.number_padding.get(), 6))
        fps = max(1, safe_int(self.fps.get(), 3))
        metadata = {
            'title': self.video_title.get().strip(),
            'artist': self.video_artist.get().strip(),
            'comment': self.video_comment.get().strip(),
            'date': datetime.datetime.now().strftime('%Y-%m-%d'),
        }

        def export_task() -> None:
            try:
                output_dir.mkdir(parents=True, exist_ok=True)
                exported: list[Path] = []

                for i, idx in enumerate(indices):
                    if progress.cancelled:
                        break

                    self.root.after(0, lambda i=i: progress.update_progress(
                        i + 1, total, f"Exporting image {i + 1}...",
                        f"{i + 1} / {total}"))

                    if idx in self.image_cache and len(self.image_cache[idx]) >= IMAGE_DATA_SIZE:
                        img = decode_gb_camera_image(self.image_cache[idx], palette)
                        if scale > 1:
                            img = img.resize(
                                (img.width * scale, img.height * scale),
                                Image.NEAREST)
                        filepath = output_dir / f"{str(i).zfill(padding)}.bmp"
                        img.save(filepath, 'BMP')
                        exported.append(filepath)

                if progress.cancelled:
                    self.root.after(0, progress.complete)
                    return

                if gif_path and exported:
                    self.root.after(0, lambda: progress.update_progress(
                        total, total, "Creating GIF...", ""))
                    create_gif(exported, gif_path, fps)

                if mkv_path and exported:
                    self.root.after(0, lambda: progress.update_progress(
                        total, total, "Creating MKV...", ""))
                    create_mkv(output_dir, mkv_path, fps, metadata)

                self.root.after(0, progress.complete)
                final_path = mkv_path or gif_path or output_dir
                count = len(exported)
                self.root.after(0, lambda: self.set_status(
                    f"Exported {count} images \u2192 {final_path}"))

            except Exception as e:
                self.root.after(0, progress.complete)
                self.root.after(0, lambda: messagebox.showerror("Export Error", str(e)))
                self.root.after(0, lambda: self.set_status("Export failed"))

        thread = threading.Thread(target=export_task, daemon=True)
        thread.start()

    # dialogs
    def rescan_images(self) -> None:
        if not self.connected:
            messagebox.showwarning("Not Connected", "Connect to device first.")
            return
        if self.dump_running:
            messagebox.showwarning("Scan Running", "Cancel current scan first.")
            return
        self.start_scan()

    def erase_images(self) -> None:
        if not self.connected:
            messagebox.showwarning("Not Connected", "Connect to device first.")
            return

        bitmap = self.device.read_bitmap()
        count = len(self.device.get_filled_slots(bitmap)) if bitmap else 0

        msg = (
            f"Device has {count} images stored.\n\n"
            "This will erase the allocation bitmap, marking all slots as empty.\n"
            "Image data is NOT zeroed, just marked as available for overwrite.\n\n"
            "This cannot be undone. Continue?"
        )
        if not messagebox.askyesno("Erase All Images", msg, icon='warning'):
            return

        self.set_status("Erasing bitmap...")
        self.root.update_idletasks()

        if self.device.erase_bitmap():
            # Verify by reading back
            bitmap = self.device.read_bitmap()
            remaining = len(self.device.get_filled_slots(bitmap)) if bitmap else -1
            if remaining == 0:
                self.set_status("Erase complete. Device is empty.")
                messagebox.showinfo("Erase Complete",
                    "Bitmap erased successfully. Device reports 0 images.")
            elif remaining > 0:
                self.set_status(f"Erase may have failed. {remaining} images still reported.")
                messagebox.showwarning("Erase Incomplete",
                    f"Write appeared to succeed but device still reports "
                    f"{remaining} images.\n\n"
                    "The device may not support bitmap writes via this protocol.")
            else:
                self.set_status("Erase sent but verification read failed.")
                messagebox.showwarning("Erase Unverified",
                    "Erase command sent but could not read bitmap to verify.")
        else:
            self.set_status("Erase may have failed.")
            messagebox.showerror("Erase Failed",
                "Erase may have failed. Check device connection.")

    def show_device_info(self) -> None:
        if not self.connected:
            messagebox.showwarning("Not Connected", "Connect to device first.")
            return

        lines = [
            f"Port: {self.device.port}",
            f"Protocol: ASCII, 1,000,000 baud",
            "",
            f"Valid images loaded: {len(self.valid_indices)}",
            f"Max capacity: 18,720 images",
        ]
        messagebox.showinfo("Device Information", "\n".join(lines))

    def show_user_guide(self) -> None:
        guide = """\
GETTING STARTED

1. Connect your PicNRec AIO via USB-A to USB-C cable
2. Click "Connect": device info appears immediately
3. Click "Scan" to load images
4. Cancel anytime to stop the scan
5. Browse with arrow keys or the navigator


PALETTES

Click a palette swatch to change colors:
  DMG (Original Game Boy green)
  Pocket (Game Boy Pocket grayscale)
  Grayscale (pure black/white)
  Inverted (negative image)
  Harsh (high contrast, no grays)

Use the Color Adjustment sliders for custom colors.


EXPORTING

The Export fields use gallery positions (1-indexed).
If you have 13 valid images, enter 1-13 for all.

  BMP (individual image files)
  GIF (animated GIF)
  MKV (video file)


TROUBLESHOOTING

If images fail to load:
  The CH340 chip can be flaky: try disconnect/reconnect
  Each slot is retried up to 3 times automatically
  Use Cancel to stop, then Re-scan to retry"""

        dialog = tk.Toplevel(self.root)
        dialog.title("User Guide")
        dialog.transient(self.root)
        dialog.configure(bg=self.theme.colors['bg'])

        frame = ttk.Frame(dialog, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        text_frame = ttk.Frame(frame)
        text_frame.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        text = tk.Text(text_frame, wrap=tk.WORD, width=55, height=28,
                       bg=self.theme.colors['bg_secondary'],
                       fg=self.theme.colors['fg'],
                       font=('TkDefaultFont', 10),
                       relief='flat', padx=15, pady=15,
                       yscrollcommand=scrollbar.set)
        text.insert('1.0', guide)
        text.configure(state='disabled')
        text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=text.yview)

        ttk.Button(frame, text="Close",
                   command=dialog.destroy).pack(pady=(15, 0))

        for key in ('<Return>', '<KP_Enter>', '<Escape>'):
            dialog.bind(key, lambda e: dialog.destroy())

        dialog.update_idletasks()
        dialog.geometry("550x550")
        x = self.root.winfo_x() + (self.root.winfo_width() - 550) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 550) // 2
        dialog.geometry(f"+{x}+{y}")

    def show_shortcuts(self) -> None:
        shortcuts = """\
Navigation:
  \u2190  Previous image
  \u2192  Next image
  Home  First image
  End  Last image

File:
  Ctrl+O  Connect to device
  Ctrl+W  Disconnect
  Ctrl+E  Export images
  Ctrl+Q  Quit"""

        dialog = tk.Toplevel(self.root)
        dialog.title("Keyboard Shortcuts")
        dialog.transient(self.root)
        dialog.configure(bg=self.theme.colors['bg'])

        frame = ttk.Frame(dialog, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        text = tk.Text(frame, wrap=tk.NONE, width=45, height=18,
                       bg=self.theme.colors['bg_secondary'],
                       fg=self.theme.colors['fg'],
                       font=('TkMonoFont', 10),
                       relief='flat', padx=10, pady=10)
        text.insert('1.0', shortcuts)
        text.configure(state='disabled')
        text.pack()

        ttk.Button(frame, text="Close",
                   command=dialog.destroy).pack(pady=(15, 0))

        for key in ('<Return>', '<KP_Enter>', '<Escape>'):
            dialog.bind(key, lambda e: dialog.destroy())

        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - dialog.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{x}+{y}")

    def show_about(self) -> None:
        about_text = """\
A reimplementation of the PicNRec
interface for more platforms.

Original hardware and Windows software by
insideGadgets (Alex). (I think?)"""

        dialog = tk.Toplevel(self.root)
        dialog.title("About This")
        dialog.transient(self.root)
        dialog.resizable(False, False)
        dialog.configure(bg=self.theme.colors['bg'])

        frame = ttk.Frame(dialog, padding=30)
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="PicNRec Tk",
                  font=('TkDefaultFont', 16, 'bold')).pack()
        ttk.Label(frame, text="PicNRec for not-just-Windows",
                  style='Secondary.TLabel').pack(pady=(0, 15))

        text = tk.Text(frame, wrap=tk.WORD, width=45, height=7,
                       bg=self.theme.colors['bg'],
                       fg=self.theme.colors['fg'],
                       font=('TkDefaultFont', 10),
                       relief='flat', padx=5, pady=5)
        text.insert('1.0', about_text)
        text.configure(state='disabled')
        text.pack()

        ttk.Button(frame, text="Close",
                   command=dialog.destroy).pack(pady=(15, 0))

        for key in ('<Return>', '<KP_Enter>', '<Escape>'):
            dialog.bind(key, lambda e: dialog.destroy())

        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - dialog.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{x}+{y}")


def main() -> int:
    root = tk.Tk()
    PicNRecGUI(root)
    root.mainloop()
    return 0


if __name__ == '__main__':
    import sys
    sys.exit(main())
