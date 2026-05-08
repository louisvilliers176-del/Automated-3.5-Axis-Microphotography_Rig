import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import serial
import serial.tools.list_ports
import threading
import time
import os
import json
import math
import base64
import io
import sys
import csv
import subprocess


# ==============================================================================
# 🎨 DARK THEME PALETTE
# ==============================================================================
T = {
    'bg':      '#0A0E17',   # root background
    'panel':   '#111827',   # panel / LabelFrame background
    'widget':  '#1F2937',   # entries, text fields
    'hover':   '#374151',   # hover state
    'border':  '#1E3A5F',   # subtle borders / separators
    'accent':  '#38BDF8',   # primary accent (sky blue)
    'accent2': '#0EA5E9',   # secondary accent (darker)
    'success': '#22C55E',   # connected / ok
    'error':   '#EF4444',   # errors
    'warning': '#F59E0B',   # warnings
    'info':    '#60A5FA',   # info messages
    'text':    '#F1F5F9',   # primary text
    'text2':   '#94A3B8',   # secondary / dim text
    'text3':   '#475569',   # very dim text
}


def get_save_path(filename):
    if getattr(sys, 'frozen', False):
        datadir = os.path.dirname(sys.executable)
    else:
        datadir = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(datadir, filename)


def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
    return os.path.join(base_path, relative_path)


try:
    from PIL import Image, ImageTk, ImageGrab
except ImportError:
    Image = None
    ImageTk = None
    ImageGrab = None
    messagebox.showerror("Dependency Error", "Please install Pillow: 'pip install Pillow'")
    exit()


def play_sound():
    try:
        if os.name == 'nt':
            import winsound
            winsound.PlaySound("SystemHand", winsound.SND_ALIAS)
        else:
            print('\a', end='', flush=True)
    except Exception:
        print("INFO: End of sequence sound.")


# --- BASE64 IMAGE UTILITIES ---
def encode_image_to_b64(img):
    if not img:
        return ""
    try:
        if img.mode != 'RGB':
            img = img.convert('RGB')
        img.thumbnail((150, 150))
        buffered = io.BytesIO()
        img.save(buffered, format="JPEG", quality=80)
        return base64.b64encode(buffered.getvalue()).decode("utf-8")
    except Exception:
        return ""


def decode_b64_to_image(b64_str):
    if not b64_str:
        return None
    try:
        img_data = base64.b64decode(b64_str)
        return Image.open(io.BytesIO(img_data))
    except Exception:
        return None


# ==============================================================================
# 🧠 OPTICAL MODULE (Mathematics)
# ==============================================================================
class OpticalCalculator:
    LIGHT_WAVELENGTH_MM = 0.00055  # 550nm

    @staticmethod
    def get_focal_length(nominal_mag, obj_tube_length):
        return obj_tube_length / (nominal_mag + 1.0)

    @staticmethod
    def calculate_working_distance(m_pos, f, start_tube_length):
        i = start_tube_length + m_pos
        if i <= f:
            return None
        return (i * f) / (i - f)

    @classmethod
    def calculate_steps(cls, nom_mag, na, obj_tube, start_tube, current_m_pos, cam_w, cam_h, coc, prec_x, prec_y,
                        prec_z, ov_lat, ov_foc):
        f = cls.get_focal_length(nom_mag, obj_tube)
        current_i = start_tube + current_m_pos
        if current_i <= f:
            raise ValueError("M-Axis extension too short for this lens to focus.")

        M = (current_i - f) / f
        fov_y = cam_w / M
        fov_z = cam_h / M

        raw_step_y = fov_y * (1.0 - ov_lat)
        raw_step_z = fov_z * (1.0 - ov_lat)

        dof = (cls.LIGHT_WAVELENGTH_MM / (na * na)) + (coc / (M * na))
        raw_step_x = dof * (1.0 - ov_foc)

        def quantize(raw_val, prec):
            if prec <= 0:
                return raw_val
            # Number of decimal places derived from machine precision (never less than 6 for fine steps)
            if prec >= 1.0:
                ndigits = 4
            else:
                ndigits = max(6, -int(math.floor(math.log10(prec))))
            # Always floor: the actual step is always ≤ the raw step,
            # which ensures the effective overlap is always ≥ the requested overlap
            q = math.floor(raw_val / prec) * prec
            if q < prec:
                q = prec
            return round(q, ndigits)

        return {
            "M": M, "dof": dof, "fov_y": fov_y, "fov_z": fov_z,
            "step_x": quantize(raw_step_x, prec_x),
            "step_y": quantize(raw_step_y, prec_y),
            "step_z": quantize(raw_step_z, prec_z)
        }


# ==============================================================================
# PROFILE DIALOGS
# ==============================================================================
class LensProfileDialog(tk.Toplevel):
    def __init__(self, parent, current_profile=None):
        super().__init__(parent)
        self.title("Optical Lens Profile")
        self.geometry("500x420")
        self.transient(parent)
        self.grab_set()
        self.result = None
        self.current_img = None
        self.configure(bg=T['panel'])

        left_frame = ttk.Frame(self)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        right_frame = ttk.Frame(self)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=10, pady=10)

        ttk.Label(left_frame, text="Profile Name:").grid(row=0, column=0, pady=5, sticky="w")
        self.name_var = tk.StringVar(value=current_profile.get("name", "New Lens") if current_profile else "New Lens")
        ttk.Entry(left_frame, textvariable=self.name_var, width=20).grid(row=0, column=1, pady=5, sticky="ew")

        ttk.Label(left_frame, text="Nominal Mag (e.g., 10):").grid(row=1, column=0, pady=5, sticky="w")
        self.nom_mag_var = tk.StringVar(
            value=str(current_profile.get("nominal_mag", "10.0")) if current_profile else "10.0")
        ttk.Entry(left_frame, textvariable=self.nom_mag_var).grid(row=1, column=1, pady=5, sticky="ew")

        ttk.Label(left_frame, text="Numerical Aperture (NA):").grid(row=2, column=0, pady=5, sticky="w")
        self.na_var = tk.StringVar(value=str(current_profile.get("na", "0.25")) if current_profile else "0.25")
        ttk.Entry(left_frame, textvariable=self.na_var).grid(row=2, column=1, pady=5, sticky="ew")

        ttk.Label(left_frame, text="Objective Tube (mm):").grid(row=3, column=0, pady=5, sticky="w")
        self.obj_tube_var = tk.StringVar(
            value=str(current_profile.get("obj_tube_length", "160.0")) if current_profile else "160.0")
        ttk.Entry(left_frame, textvariable=self.obj_tube_var).grid(row=3, column=1, pady=5, sticky="ew")

        ttk.Label(left_frame, text="Current Physical Length:").grid(row=4, column=0, pady=5, sticky="w")
        self.start_tube_var = tk.StringVar(
            value=str(current_profile.get("start_tube_length", "160.0")) if current_profile else "160.0")
        ttk.Entry(left_frame, textvariable=self.start_tube_var).grid(row=4, column=1, pady=5, sticky="ew")

        self.invert_x_var = tk.BooleanVar(value=current_profile.get("invert_x", False) if current_profile else False)
        ttk.Checkbutton(left_frame, text="Invert X Axis Comp.", variable=self.invert_x_var).grid(row=5, column=0,
                                                                                                 columnspan=2, pady=5,
                                                                                                 sticky="w")

        self.img_label = tk.Label(right_frame, text="No Image", relief="flat",
                                   background=T['widget'], foreground=T['text2'],
                                   width=20, font=('Segoe UI', 9))
        self.img_label.pack(ipady=40, fill=tk.X, pady=5)

        ttk.Button(right_frame, text="Load from File", command=self.load_file).pack(fill=tk.X, pady=2)
        ttk.Button(right_frame, text="Paste Clipboard", command=self.load_clipboard).pack(fill=tk.X, pady=2)

        if current_profile and current_profile.get("image_b64"):
            img = decode_b64_to_image(current_profile["image_b64"])
            if img:
                self.set_image(img)

        btn_frame = ttk.Frame(left_frame)
        btn_frame.grid(row=6, column=0, columnspan=2, pady=15)
        ttk.Button(btn_frame, text="Save", command=self.save).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.destroy).pack(side=tk.LEFT, padx=5)

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Images", "*.png *.jpg *.jpeg")])
        if path:
            try:
                self.set_image(Image.open(path))
            except Exception:
                pass

    def load_clipboard(self):
        try:
            img = ImageGrab.grabclipboard()
            if img:
                self.set_image(img)
            else:
                messagebox.showinfo("Clipboard", "No image found in clipboard.")
        except Exception:
            messagebox.showerror("Clipboard Error", "Failed to grab clipboard.")

    def set_image(self, img):
        self.current_img = img
        img_thumb = img.copy()
        img_thumb.thumbnail((120, 120))
        self.tk_img = ImageTk.PhotoImage(img_thumb)
        self.img_label['image'] = self.tk_img
        self.img_label['text'] = ""

    def save(self):
        try:
            nom_mag = float(self.nom_mag_var.get())
            na = float(self.na_var.get())
            obj_tube = float(self.obj_tube_var.get())
            start_tube = float(self.start_tube_var.get())
            if nom_mag <= 0 or obj_tube <= 0 or start_tube <= 0 or na <= 0:
                messagebox.showerror("Error", "Values must be > 0.")
                return
            self.result = {
                "name": self.name_var.get(), "nominal_mag": nom_mag, "na": na,
                "obj_tube_length": obj_tube, "start_tube_length": start_tube,
                "invert_x": self.invert_x_var.get(),
                "image_b64": encode_image_to_b64(self.current_img)
            }
            self.destroy()
        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter valid numbers only.")


class CameraProfileDialog(tk.Toplevel):
    def __init__(self, parent, current_profile=None):
        super().__init__(parent)
        self.title("Camera Sensor Profile")
        self.geometry("500x320")
        self.transient(parent)
        self.grab_set()
        self.result = None
        self.current_img = None
        self.configure(bg=T['panel'])

        left_frame = ttk.Frame(self)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        right_frame = ttk.Frame(self)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=10, pady=10)

        ttk.Label(left_frame, text="Camera Name:").grid(row=0, column=0, pady=5, sticky="w")
        self.name_var = tk.StringVar(
            value=current_profile.get("name", "Full Frame DSLR") if current_profile else "Full Frame DSLR")
        ttk.Entry(left_frame, textvariable=self.name_var, width=20).grid(row=0, column=1, pady=5, sticky="ew")

        ttk.Label(left_frame, text="Sensor Width (mm):").grid(row=1, column=0, pady=5, sticky="w")
        self.w_var = tk.StringVar(value=str(current_profile.get("width", "36.0")) if current_profile else "36.0")
        ttk.Entry(left_frame, textvariable=self.w_var).grid(row=1, column=1, pady=5, sticky="ew")

        ttk.Label(left_frame, text="Sensor Height (mm):").grid(row=2, column=0, pady=5, sticky="w")
        self.h_var = tk.StringVar(value=str(current_profile.get("height", "24.0")) if current_profile else "24.0")
        ttk.Entry(left_frame, textvariable=self.h_var).grid(row=2, column=1, pady=5, sticky="ew")

        ttk.Label(left_frame, text="Circle of Confusion (CoC):").grid(row=3, column=0, pady=5, sticky="w")
        self.coc_var = tk.StringVar(value=str(current_profile.get("coc", "0.03")) if current_profile else "0.03")
        ttk.Entry(left_frame, textvariable=self.coc_var).grid(row=3, column=1, pady=5, sticky="ew")

        self.img_label = tk.Label(right_frame, text="No Image", relief="flat",
                                   background=T['widget'], foreground=T['text2'],
                                   width=20, font=('Segoe UI', 9))
        self.img_label.pack(ipady=40, fill=tk.X, pady=5)

        ttk.Button(right_frame, text="Load from File", command=self.load_file).pack(fill=tk.X, pady=2)
        ttk.Button(right_frame, text="Paste Clipboard", command=self.load_clipboard).pack(fill=tk.X, pady=2)

        if current_profile and current_profile.get("image_b64"):
            img = decode_b64_to_image(current_profile["image_b64"])
            if img:
                self.set_image(img)

        btn_frame = ttk.Frame(left_frame)
        btn_frame.grid(row=5, column=0, columnspan=2, pady=15)
        ttk.Button(btn_frame, text="Save", command=self.save).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.destroy).pack(side=tk.LEFT, padx=5)

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Images", "*.png *.jpg *.jpeg")])
        if path:
            try:
                self.set_image(Image.open(path))
            except Exception:
                pass

    def load_clipboard(self):
        try:
            img = ImageGrab.grabclipboard()
            if img:
                self.set_image(img)
            else:
                messagebox.showinfo("Clipboard", "No image found in clipboard.")
        except Exception:
            messagebox.showerror("Clipboard Error", "Failed to grab clipboard.")

    def set_image(self, img):
        self.current_img = img
        img_thumb = img.copy()
        img_thumb.thumbnail((120, 120))
        self.tk_img = ImageTk.PhotoImage(img_thumb)
        self.img_label['image'] = self.tk_img
        self.img_label['text'] = ""

    def save(self):
        try:
            self.result = {
                "name": self.name_var.get(),
                "width": float(self.w_var.get()),
                "height": float(self.h_var.get()),
                "coc": float(self.coc_var.get()),
                "image_b64": encode_image_to_b64(self.current_img)
            }
            self.destroy()
        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter valid numbers.")


class MachineProfileDialog(tk.Toplevel):
    def __init__(self, parent, current_profile=None):
        super().__init__(parent)
        self.title("CNC Machine Profile")
        self.geometry("550x400")
        self.transient(parent)
        self.grab_set()
        self.result = None
        self.current_img = None
        self.configure(bg=T['panel'])

        left_frame = ttk.Frame(self)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        right_frame = ttk.Frame(self)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=10, pady=10)

        ttk.Label(left_frame, text="Machine Name:").grid(row=0, column=0, pady=5, sticky="w")
        self.name_var = tk.StringVar(value=current_profile.get("name", "My CNC") if current_profile else "My CNC")
        ttk.Entry(left_frame, textvariable=self.name_var, width=20).grid(row=0, column=1, columnspan=2, pady=5,
                                                                         sticky="ew")

        ttk.Label(left_frame, text="Axis", font=("Segoe UI", 9, "bold")).grid(row=1, column=0, pady=5, sticky="w")
        ttk.Label(left_frame, text="Travel (mm)", font=("Segoe UI", 9, "bold")).grid(row=1, column=1, pady=5)
        ttk.Label(left_frame, text="Precision (mm)", font=("Segoe UI", 9, "bold")).grid(row=1, column=2, pady=5)

        axes = ['x', 'y', 'z', 'm']
        self.t_vars = {}
        self.p_vars = {}
        defaults_t = {'x': '300.0', 'y': '300.0', 'z': '100.0', 'm': '200.0'}
        defaults_p = {'x': '0.01', 'y': '0.01', 'z': '0.001', 'm': '0.01'}

        for i, ax in enumerate(axes):
            ttk.Label(left_frame, text=f"Axis {ax.upper()}:").grid(row=2 + i, column=0, pady=2, sticky="w")
            t_val = current_profile.get(f"travel_{ax}", defaults_t[ax]) if current_profile else defaults_t[ax]
            self.t_vars[ax] = tk.StringVar(value=str(t_val))
            ttk.Entry(left_frame, textvariable=self.t_vars[ax], width=8).grid(row=2 + i, column=1, padx=5, pady=2)

            p_val = current_profile.get(f"prec_{ax}", defaults_p[ax]) if current_profile else defaults_p[ax]
            self.p_vars[ax] = tk.StringVar(value=str(p_val))
            ttk.Entry(left_frame, textvariable=self.p_vars[ax], width=8).grid(row=2 + i, column=2, padx=5, pady=2)

        self.img_label = tk.Label(right_frame, text="No Image", relief="flat",
                                   background=T['widget'], foreground=T['text2'],
                                   width=20, font=('Segoe UI', 9))
        self.img_label.pack(ipady=40, fill=tk.X, pady=5)

        ttk.Button(right_frame, text="Load from File", command=self.load_file).pack(fill=tk.X, pady=2)
        ttk.Button(right_frame, text="Paste Clipboard", command=self.load_clipboard).pack(fill=tk.X, pady=2)

        if current_profile and current_profile.get("image_b64"):
            img = decode_b64_to_image(current_profile["image_b64"])
            if img:
                self.set_image(img)

        btn_frame = ttk.Frame(left_frame)
        btn_frame.grid(row=7, column=0, columnspan=3, pady=15)
        ttk.Button(btn_frame, text="Save", command=self.save).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.destroy).pack(side=tk.LEFT, padx=5)

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Images", "*.png *.jpg *.jpeg")])
        if path:
            try:
                self.set_image(Image.open(path))
            except Exception:
                pass

    def load_clipboard(self):
        try:
            img = ImageGrab.grabclipboard()
            if img:
                self.set_image(img)
            else:
                messagebox.showinfo("Clipboard", "No image found in clipboard.")
        except Exception:
            messagebox.showerror("Clipboard Error", "Failed to grab clipboard.")

    def set_image(self, img):
        self.current_img = img
        img_thumb = img.copy()
        img_thumb.thumbnail((120, 120))
        self.tk_img = ImageTk.PhotoImage(img_thumb)
        self.img_label['image'] = self.tk_img
        self.img_label['text'] = ""

    def save(self):
        try:
            self.result = {
                "name": self.name_var.get(),
                "travel_x": float(self.t_vars['x'].get()),
                "travel_y": float(self.t_vars['y'].get()),
                "travel_z": float(self.t_vars['z'].get()),
                "travel_m": float(self.t_vars['m'].get()),
                "prec_x": float(self.p_vars['x'].get()),
                "prec_y": float(self.p_vars['y'].get()),
                "prec_z": float(self.p_vars['z'].get()),
                "prec_m": float(self.p_vars['m'].get()),
                "image_b64": encode_image_to_b64(self.current_img)
            }
            self.destroy()
        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter valid numerical values.")


# ==============================================================================
# 🎛️ MAIN APPLICATION
# ==============================================================================
class CNCApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CMCS — CNC Microphotography Control Software")
        self._setup_theme()
        try:
            self.root.state('zoomed')
        except Exception:
            self.root.attributes('-fullscreen', False)

        self.cnc = None
        self.connected_cnc = False
        self.current_pos = {"x": 0.0, "y": 0.0, "z": 0.0, "a": 0.0}
        self._is_polling = False
        self.sequence_is_paused_by_error = False

        self.user_requested_pause_event = threading.Event()
        self.sequence_pause_event = threading.Event()
        self.sequence_resume_event = threading.Event()
        self.stop_sequence_flag = threading.Event()
        self.skip_line_event = threading.Event()
        self.sequence_running = False
        self._serial_lock = threading.Lock()

        self.config_file = get_save_path("config.json")
        self.lens_profiles = {}
        self.active_lens_var = tk.StringVar()
        self.camera_profiles = {}
        self.active_camera_var = tk.StringVar()
        self.machine_profiles = {}
        self.active_machine_var = tk.StringVar()

        self.com_port_var = tk.StringVar()
        self.start_pos_vars = {axis: tk.StringVar(value="0.0") for axis in "xyz"}
        self.end_pos_vars = {axis: tk.StringVar(value="0.0") for axis in "xyz"}
        self.step_seq_vars = {axis: tk.StringVar(value="1.0") for axis in "xyz"}
        self.delay_var = tk.StringVar(value="0.5")
        self.speed_var = tk.StringVar(value="1000")
        self.step_var = tk.StringVar(value="1")
        self.step_m_var = tk.StringVar(value="1")
        self.progress_var = tk.DoubleVar(value=0.0)

        self.overlap_lat_var = tk.StringVar(value="40")
        self.overlap_focus_var = tk.StringVar(value="20")

        # Variables for the Log module
        self.log_enabled_var = tk.BooleanVar(value=True)
        default_log_path = os.path.join(os.path.expanduser("~"), "Desktop", "CNC_Logs")
        self.log_dir_var = tk.StringVar(value=default_log_path)

        self.drag_source = None
        self.drag_win = None
        self.drop_target = None
        self.drop_indicator = tk.Frame(self.root, bg=T['accent'], height=3)
        self.panels = {}
        self.layout = {"left": ["cnc", "trigger", "log_panel", "mag"], "center": ["hw", "seq", "vis"],
                       "right": ["grbl", "hist"]}
        self.saved_sashes = {}
        self.seq_info = None

        self._load_config()

        top_frame = ttk.Frame(root, padding="5")
        top_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.main_pw = ttk.PanedWindow(top_frame, orient=tk.HORIZONTAL)
        self.main_pw.pack(fill=tk.BOTH, expand=True)
        self.left_pw = ttk.PanedWindow(self.main_pw, orient=tk.VERTICAL)
        self.center_pw = ttk.PanedWindow(self.main_pw, orient=tk.VERTICAL)
        self.right_pw = ttk.PanedWindow(self.main_pw, orient=tk.VERTICAL)

        self.main_pw.add(self.left_pw, weight=1)
        self.main_pw.add(self.center_pw, weight=1)
        self.main_pw.add(self.right_pw, weight=1)

        self.panels["cnc"] = ttk.LabelFrame(self.root, text="CNC Link & Movement")
        self._create_cnc_control_widgets(self.panels["cnc"])
        self.panels["trigger"] = ttk.LabelFrame(self.root, text="Camera Control")
        self._create_trigger_widgets(self.panels["trigger"])

        # New Logging Panel
        self.panels["log_panel"] = ttk.LabelFrame(self.root, text="SD-Card Sync (Logs)")
        self._create_logging_widgets(self.panels["log_panel"])

        self.panels["hw"] = ttk.LabelFrame(self.root, text="Hardware Setup")
        self._create_hw_widgets(self.panels["hw"])
        self.panels["mag"] = ttk.LabelFrame(self.root, text="Magnification")
        self._create_magnification_widgets(self.panels["mag"])
        self.panels["seq"] = ttk.LabelFrame(self.root, text="Focus Stacking Sequence")
        self._create_sequence_control_widgets(self.panels["seq"])
        self.panels["grbl"] = ttk.LabelFrame(self.root, text="GRBL Console")
        self._create_grbl_console_widgets(self.panels["grbl"])
        self.panels["hist"] = ttk.LabelFrame(self.root, text="Sequence Log")
        self._create_history_widgets(self.panels["hist"])
        self.panels["vis"] = ttk.LabelFrame(self.root, text="3D Sequence View")
        self._create_3d_view_widgets(self.panels["vis"])

        for pid, panel in self.panels.items():
            self._bind_drag(panel)

        self._apply_layout()
        self.root.after(200, self._restore_sashes)

        self.update_com_ports(select_last=True)
        self._update_hw_combos()
        self.update_position_display()
        self.update_sequence_buttons_state()

        threading.Thread(target=self._auto_connect_worker, daemon=True).start()

    # --- THEME SETUP ---
    def _setup_theme(self):
        """Configure a modern dark instrument-panel theme for all ttk widgets."""
        style = ttk.Style(self.root)
        style.theme_use('clam')

        BG     = T['bg']
        PANEL  = T['panel']
        WIDGET = T['widget']
        HOVER  = T['hover']
        BORDER = T['border']
        ACCENT = T['accent']
        TEXT   = T['text']
        TEXT2  = T['text2']
        TEXT3  = T['text3']

        FONT      = ('Segoe UI', 9)
        FONT_B    = ('Segoe UI', 9, 'bold')
        FONT_SM   = ('Segoe UI', 8)
        FONT_MONO = ('Consolas', 9)

        self.root.configure(bg=BG)

        # --- Base ---
        style.configure('.',
            background=PANEL, foreground=TEXT,
            font=FONT, borderwidth=0,
            troughcolor=WIDGET, focuscolor=ACCENT,
            selectbackground=BORDER, selectforeground=TEXT)

        # --- Frame / PanedWindow ---
        style.configure('TFrame', background=PANEL)
        style.configure('TPanedwindow', background=BG)
        style.configure('Sash', background=BORDER, sashthickness=5, sashpad=2)

        # --- LabelFrame (panels) ---
        style.configure('TLabelframe',
            background=PANEL, foreground=ACCENT,
            bordercolor=BORDER, relief='flat', borderwidth=1)
        style.configure('TLabelframe.Label',
            background=PANEL, foreground=ACCENT,
            font=FONT_B, padding=(4, 2))

        # --- Labels ---
        style.configure('TLabel', background=PANEL, foreground=TEXT, font=FONT)

        # Specialized label styles
        style.configure('Coord.TLabel',
            background=PANEL, foreground=ACCENT,
            font=('Consolas', 12, 'bold'))
        style.configure('CoordKey.TLabel',
            background=PANEL, foreground=TEXT2, font=FONT)
        style.configure('Mag.TLabel',
            background=PANEL, foreground=ACCENT, font=FONT_B)
        style.configure('Dim.TLabel',
            background=PANEL, foreground=TEXT3, font=FONT_SM)
        style.configure('Img.TLabel',
            background=WIDGET, foreground=TEXT2,
            relief='flat', borderwidth=1)

        # --- Entry ---
        style.configure('TEntry',
            fieldbackground=WIDGET, foreground=TEXT,
            bordercolor=BORDER, insertcolor=ACCENT,
            selectbackground=BORDER, selectforeground=TEXT,
            font=FONT, relief='flat', padding=4)
        style.map('TEntry',
            fieldbackground=[('focus', HOVER)],
            bordercolor=[('focus', ACCENT)])

        # --- Combobox ---
        style.configure('TCombobox',
            fieldbackground=WIDGET, foreground=TEXT,
            background=WIDGET, arrowcolor=ACCENT,
            bordercolor=BORDER, selectbackground=BORDER,
            selectforeground=TEXT, font=FONT, padding=3)
        style.map('TCombobox',
            fieldbackground=[('readonly', WIDGET), ('focus', HOVER)],
            selectbackground=[('readonly', WIDGET)],
            foreground=[('readonly', TEXT)],
            arrowcolor=[('disabled', TEXT3), ('pressed', ACCENT)])

        # --- Button ---
        style.configure('TButton',
            background=WIDGET, foreground=TEXT,
            bordercolor=BORDER, relief='flat',
            padding=(10, 5), font=FONT,
            focuscolor=ACCENT, anchor='center')
        style.map('TButton',
            background=[('active', HOVER), ('pressed', BORDER), ('disabled', PANEL)],
            foreground=[('active', ACCENT), ('disabled', TEXT3)],
            bordercolor=[('active', ACCENT)])

        # Accent (primary action) button
        style.configure('Accent.TButton',
            background=ACCENT, foreground=BG,
            font=FONT_B, padding=(10, 5), relief='flat')
        style.map('Accent.TButton',
            background=[('active', '#7DD3FC'), ('pressed', '#0284C7')])

        # Danger (stop) button
        style.configure('Danger.TButton',
            background='#7F1D1D', foreground=T['error'],
            font=FONT, padding=(10, 5), relief='flat')
        style.map('Danger.TButton',
            background=[('active', '#991B1B'), ('disabled', PANEL)],
            foreground=[('disabled', TEXT3)])

        # --- Checkbutton ---
        style.configure('TCheckbutton',
            background=PANEL, foreground=TEXT,
            font=FONT, focuscolor=ACCENT)
        style.map('TCheckbutton',
            background=[('active', PANEL)],
            foreground=[('active', ACCENT)])

        # --- Progressbar ---
        style.configure('TProgressbar',
            troughcolor=WIDGET, background=ACCENT,
            bordercolor=BORDER, lightcolor=ACCENT,
            darkcolor=ACCENT, thickness=6)

        # --- Scrollbar ---
        style.configure('TScrollbar',
            background=HOVER, troughcolor=WIDGET,
            arrowcolor=TEXT2, bordercolor=PANEL,
            darkcolor=HOVER, lightcolor=HOVER)
        style.map('TScrollbar',
            background=[('active', TEXT2)],
            arrowcolor=[('active', ACCENT)])

        # --- Separator ---
        style.configure('TSeparator', background=BORDER)

    # --- WIDGETS CREATION ---
    def _create_logging_widgets(self, parent):
        ttk.Checkbutton(parent, text="Enable XYZ Position Logging (CSV)", variable=self.log_enabled_var).pack(
            anchor="w", padx=5, pady=2)

        path_frame = ttk.Frame(parent)
        path_frame.pack(fill=tk.X, padx=5, pady=2)
        ttk.Entry(path_frame, textvariable=self.log_dir_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(path_frame, text="Browse", width=8, command=self._browse_log_dir).pack(side=tk.LEFT)

        ttk.Button(parent, text="Open Logs Folder", command=self._open_log_dir).pack(fill=tk.X, padx=5, pady=(5, 10))

    def _browse_log_dir(self):
        dir_path = filedialog.askdirectory(initialdir=self.log_dir_var.get())
        if dir_path:
            self.log_dir_var.set(dir_path)

    def _open_log_dir(self):
        path = self.log_dir_var.get()
        if not os.path.exists(path):
            try:
                os.makedirs(path)
            except Exception as e:
                messagebox.showerror("Error", f"Cannot create folder: {e}")
                return
        try:
            if os.name == 'nt':
                os.startfile(path)
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', path])
            else:
                subprocess.Popen(['xdg-open', path])
        except Exception as e:
            messagebox.showerror("Error", f"Cannot open folder: {e}")

    def _create_trigger_widgets(self, parent):
        self.manual_trigger_button = ttk.Button(parent, text="📸 Manual Camera Trigger",
                                                command=self.trigger_camera_manually)
        self.manual_trigger_button.pack(fill=tk.BOTH, expand=True, padx=10, pady=10, ipadx=10, ipady=10)

    def _create_hw_widgets(self, parent):
        parent.columnconfigure(0, weight=1)
        parent.columnconfigure(1, weight=1)
        parent.columnconfigure(2, weight=1)

        # Lens
        lf = ttk.Frame(parent)
        lf.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
        ttk.Label(lf, text="Lens:").pack(anchor="w")
        c_lf = ttk.Frame(lf)
        c_lf.pack(fill=tk.X, pady=2)
        self.lens_combo = ttk.Combobox(c_lf, textvariable=self.active_lens_var, state="readonly")
        self.lens_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(c_lf, text="Edit", width=4, command=self.open_lens_editor).pack(side=tk.RIGHT, padx=2)
        ttk.Button(c_lf, text="✕", width=2, command=self.delete_lens_profile).pack(side=tk.RIGHT, padx=(0, 2))
        self.lens_img_label = tk.Label(lf, text="No Image",
                                        background=T['widget'], foreground=T['text2'],
                                        relief="flat", font=('Segoe UI', 9))
        self.lens_img_label.pack(fill=tk.BOTH, expand=True, pady=2)
        self.active_lens_var.trace_add("write", lambda var, index, mode: self._refresh_hw_display())

        # Camera
        cf = ttk.Frame(parent)
        cf.grid(row=0, column=1, sticky="nsew", padx=2, pady=2)
        ttk.Label(cf, text="Camera:").pack(anchor="w")
        c_cf = ttk.Frame(cf)
        c_cf.pack(fill=tk.X, pady=2)
        self.cam_combo = ttk.Combobox(c_cf, textvariable=self.active_camera_var, state="readonly")
        self.cam_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(c_cf, text="Edit", width=4, command=self.open_camera_editor).pack(side=tk.RIGHT, padx=2)
        ttk.Button(c_cf, text="✕", width=2, command=self.delete_camera_profile).pack(side=tk.RIGHT, padx=(0, 2))
        self.cam_img_label = tk.Label(cf, text="No Image",
                                       background=T['widget'], foreground=T['text2'],
                                       relief="flat", font=('Segoe UI', 9))
        self.cam_img_label.pack(fill=tk.BOTH, expand=True, pady=2)
        self.active_camera_var.trace_add("write", lambda var, index, mode: self._refresh_hw_display())

        # Machine
        rf = ttk.Frame(parent)
        rf.grid(row=0, column=2, sticky="nsew", padx=2, pady=2)
        ttk.Label(rf, text="Machine:").pack(anchor="w")
        c_rf = ttk.Frame(rf)
        c_rf.pack(fill=tk.X, pady=2)
        self.mach_combo = ttk.Combobox(c_rf, textvariable=self.active_machine_var, state="readonly")
        self.mach_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(c_rf, text="Edit", width=4, command=self.open_machine_editor).pack(side=tk.RIGHT, padx=2)
        ttk.Button(c_rf, text="✕", width=2, command=self.delete_machine_profile).pack(side=tk.RIGHT, padx=(0, 2))
        self.mach_img_label = tk.Label(rf, text="No Image",
                                        background=T['widget'], foreground=T['text2'],
                                        relief="flat", font=('Segoe UI', 9))
        self.mach_img_label.pack(fill=tk.BOTH, expand=True, pady=2)
        self.active_machine_var.trace_add("write", lambda var, index, mode: self._refresh_hw_display())

    def _refresh_hw_display(self):
        def update_img(prof_dict, var_name, label, tk_img_attr):
            prof = prof_dict.get(var_name.get())
            if prof and prof.get("image_b64"):
                img = decode_b64_to_image(prof["image_b64"])
                if img:
                    setattr(self, tk_img_attr, ImageTk.PhotoImage(img))
                    label.config(image=getattr(self, tk_img_attr), text="")
                    return
            label.config(image="", text="No Image")

        update_img(self.lens_profiles, self.active_lens_var, self.lens_img_label, "tk_lens_img")
        update_img(self.camera_profiles, self.active_camera_var, self.cam_img_label, "tk_cam_img")
        update_img(self.machine_profiles, self.active_machine_var, self.mach_img_label, "tk_mach_img")
        self.update_position_display()

    def open_lens_editor(self):
        curr = self.lens_profiles.get(self.active_lens_var.get())
        dialog = LensProfileDialog(self.root, curr)
        self.root.wait_window(dialog)
        if dialog.result:
            self.lens_profiles[dialog.result["name"]] = dialog.result
            self._update_hw_combos(select_lens=dialog.result["name"])
            self._save_config()

    def open_camera_editor(self):
        curr = self.camera_profiles.get(self.active_camera_var.get())
        dialog = CameraProfileDialog(self.root, curr)
        self.root.wait_window(dialog)
        if dialog.result:
            self.camera_profiles[dialog.result["name"]] = dialog.result
            self._update_hw_combos(select_cam=dialog.result["name"])
            self._save_config()

    def open_machine_editor(self):
        curr = self.machine_profiles.get(self.active_machine_var.get())
        dialog = MachineProfileDialog(self.root, curr)
        self.root.wait_window(dialog)
        if dialog.result:
            self.machine_profiles[dialog.result["name"]] = dialog.result
            self._update_hw_combos(select_mach=dialog.result["name"])
            self._save_config()

    def delete_lens_profile(self):
        name = self.active_lens_var.get()
        if not name or name not in self.lens_profiles:
            return
        if not messagebox.askyesno("Delete Profile", f"Delete lens profile '{name}'?"):
            return
        del self.lens_profiles[name]
        self._update_hw_combos()
        self._save_config()

    def delete_camera_profile(self):
        name = self.active_camera_var.get()
        if not name or name not in self.camera_profiles:
            return
        if not messagebox.askyesno("Delete Profile", f"Delete camera profile '{name}'?"):
            return
        del self.camera_profiles[name]
        self._update_hw_combos()
        self._save_config()

    def delete_machine_profile(self):
        name = self.active_machine_var.get()
        if not name or name not in self.machine_profiles:
            return
        if not messagebox.askyesno("Delete Profile", f"Delete machine profile '{name}'?"):
            return
        del self.machine_profiles[name]
        self._update_hw_combos()
        self._save_config()

    def _update_hw_combos(self, select_lens=None, select_cam=None, select_mach=None):
        names_l = list(self.lens_profiles.keys())
        if not names_l:
            names_l = ["Default 10x Lens"]
            self.lens_profiles["Default 10x Lens"] = {
                "name": "Default 10x Lens", "nominal_mag": 10.0, "na": 0.25,
                "obj_tube_length": 160.0, "start_tube_length": 160.0, "invert_x": False
            }
        self.lens_combo['values'] = names_l
        if select_lens and select_lens in names_l:
            self.active_lens_var.set(select_lens)
        elif not self.active_lens_var.get() in names_l:
            self.active_lens_var.set(names_l[0])

        names_c = list(self.camera_profiles.keys())
        if not names_c:
            names_c = ["Full Frame DSLR"]
            self.camera_profiles["Full Frame DSLR"] = {
                "name": "Full Frame DSLR", "width": 36.0, "height": 24.0, "coc": 0.03
            }
        self.cam_combo['values'] = names_c
        if select_cam and select_cam in names_c:
            self.active_camera_var.set(select_cam)
        elif not self.active_camera_var.get() in names_c:
            self.active_camera_var.set(names_c[0])

        names_m = list(self.machine_profiles.keys())
        if not names_m:
            names_m = ["Default CNC"]
            self.machine_profiles["Default CNC"] = {
                "name": "Default CNC", "travel_x": 300.0, "travel_y": 300.0,
                "travel_z": 100.0, "travel_m": 200.0, "prec_x": 0.01,
                "prec_y": 0.01, "prec_z": 0.001, "prec_m": 0.01
            }
        self.mach_combo['values'] = names_m
        if select_mach and select_mach in names_m:
            self.active_machine_var.set(select_mach)
        elif not self.active_machine_var.get() in names_m:
            self.active_machine_var.set(names_m[0])

        self._refresh_hw_display()

    # --- MAGNETIC DRAG AND DROP ---
    def _bind_drag(self, panel):
        panel.bind("<ButtonPress-1>", lambda e, p=panel: self._on_drag_start(e, p))
        panel.bind("<B1-Motion>", self._on_drag_motion)
        panel.bind("<ButtonRelease-1>", self._on_drag_release)

    def _on_drag_start(self, event, panel):
        if event.y <= 25:
            self.drag_source = panel
            self.root.config(cursor="fleur")
            self.drag_win = tk.Toplevel(self.root)
            self.drag_win.overrideredirect(True)
            self.drag_win.attributes('-alpha', 0.92)
            self.drag_win.configure(bg=T['accent'])
            lbl = tk.Label(self.drag_win, text=f" ↕  {panel.cget('text')} ",
                           font=("Segoe UI", 9, "bold"),
                           bg=T['accent'], fg=T['bg'], relief="flat", padx=10, pady=8)
            lbl.pack()
            self.drag_win.geometry(f"+{event.x_root + 15}+{event.y_root + 15}")

    def _on_drag_motion(self, event):
        if not self.drag_source:
            return
        if self.drag_win:
            self.drag_win.geometry(f"+{event.x_root + 15}+{event.y_root + 15}")

        root_x, root_y = self.root.winfo_rootx(), self.root.winfo_rooty()
        mouse_x, mouse_y = self.root.winfo_pointerxy()
        rel_x, rel_y = mouse_x - root_x, mouse_y - root_y

        center_x = self.center_pw.winfo_rootx() - root_x
        right_x = self.right_pw.winfo_rootx() - root_x

        target_pw, target_col = self.left_pw, "left"
        if rel_x > right_x and right_x > 0:
            target_pw, target_col = self.right_pw, "right"
        elif rel_x > center_x and center_x > 0:
            target_pw, target_col = self.center_pw, "center"

        panes = target_pw.panes()
        insert_index = len(panes)
        indicator_y = target_pw.winfo_rooty() - root_y + target_pw.winfo_height()

        for i, pane_name in enumerate(panes):
            try:
                pane_widget = self.root.nametowidget(pane_name)
                pane_y = pane_widget.winfo_rooty() - root_y
                if rel_y < pane_y + (pane_widget.winfo_height() / 2):
                    insert_index, indicator_y = i, pane_y - 2
                    break
            except Exception:
                pass

        self.drop_target = (target_col, insert_index)
        self.drop_indicator.place(x=target_pw.winfo_rootx() - root_x, y=indicator_y, width=target_pw.winfo_width())

    def _on_drag_release(self, event):
        if not self.drag_source:
            return
        self.root.config(cursor="")
        self.drop_indicator.place_forget()
        if self.drag_win:
            self.drag_win.destroy()
            self.drag_win = None

        if self.drop_target:
            target_col, insert_index = self.drop_target
            dragged_id = None
            for pid, pwidget in self.panels.items():
                if pwidget == self.drag_source:
                    dragged_id = pid
                    break
            if dragged_id:
                for col in ["left", "center", "right"]:
                    if dragged_id in self.layout[col]:
                        self.layout[col].remove(dragged_id)
                self.layout[target_col].insert(insert_index, dragged_id)
                self._apply_layout()
                self._save_config()

        self.drag_source, self.drop_target = None, None

    def _apply_layout(self):
        for p in self.left_pw.panes(): self.left_pw.forget(p)
        for p in self.center_pw.panes(): self.center_pw.forget(p)
        for p in self.right_pw.panes(): self.right_pw.forget(p)

        for pid in self.layout.get("left", []):
            if pid in self.panels: self.left_pw.add(self.panels[pid], weight=1)
        for pid in self.layout.get("center", []):
            if pid in self.panels: self.center_pw.add(self.panels[pid], weight=1)
        for pid in self.layout.get("right", []):
            if pid in self.panels: self.right_pw.add(self.panels[pid], weight=1)

    def _restore_sashes(self):
        if not hasattr(self, 'saved_sashes') or not self.saved_sashes:
            return
        try:
            for i, pos in enumerate(self.saved_sashes.get('main', [])): self.main_pw.sashpos(i, pos)
            for i, pos in enumerate(self.saved_sashes.get('left', [])): self.left_pw.sashpos(i, pos)
            for i, pos in enumerate(self.saved_sashes.get('center', [])): self.center_pw.sashpos(i, pos)
            for i, pos in enumerate(self.saved_sashes.get('right', [])): self.right_pw.sashpos(i, pos)
        except Exception:
            pass

    # ==============================================================================
    # 🎨 ADVANCED ISOMETRIC 3D ENGINE (X Axis = Depth)
    # ==============================================================================
    def _create_3d_view_widgets(self, parent):
        self.vis_canvas = tk.Canvas(parent, bg=T['bg'], highlightthickness=0)
        self.vis_canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.vis_canvas.bind("<Configure>", self._redraw_3d_view)

    def update_3d_preview_data(self, x_pts=None, y_pts=None, z_pts=None):
        try:
            start_pos = {axis: float(var.get()) for axis, var in self.start_pos_vars.items()}
            end_pos = {axis: float(var.get()) for axis, var in self.end_pos_vars.items()}
            steps = {axis: float(var.get()) for axis, var in self.step_seq_vars.items()}

            if not x_pts or not y_pts or not z_pts:
                x_pts = self._generate_scan_points(start_pos['x'], end_pos['x'], steps['x'])
                y_pts = self._generate_scan_points(start_pos['y'], end_pos['y'], steps['y'])
                z_pts = self._generate_scan_points(start_pos['z'], end_pos['z'], steps['z'])

            prof_l = self.lens_profiles.get(self.active_lens_var.get())
            prof_c = self.camera_profiles.get(self.active_camera_var.get())

            fov_y, fov_z = 5.0, 5.0
            if prof_l and prof_c:
                try:
                    nom_mag = float(prof_l.get("nominal_mag", 10.0))
                    obj_tube = float(prof_l.get("obj_tube_length", 160.0))
                    start_tube = float(prof_l.get("start_tube_length", 160.0))
                    cam_w = float(prof_c.get("width", 36.0))
                    cam_h = float(prof_c.get("height", 24.0))

                    f = OpticalCalculator.get_focal_length(nom_mag, obj_tube)
                    current_i = start_tube + self.current_pos["a"]
                    if current_i > f:
                        M = (current_i - f) / f
                        fov_y = cam_w / M
                        fov_z = cam_h / M
                except Exception:
                    pass

            self.seq_info = {
                'x_pts': x_pts, 'y_pts': y_pts, 'z_pts': z_pts,
                'x_min': min(x_pts), 'x_max': max(x_pts),
                'y_min': min(y_pts), 'y_max': max(y_pts),
                'z_min': min(z_pts), 'z_max': max(z_pts),
                'cx': self.current_pos['x'], 'cy': self.current_pos['y'], 'cz': self.current_pos['z'],
                'fov_y': fov_y, 'fov_z': fov_z,
                'start_x': x_pts[0] if x_pts else start_pos['x']
            }
            self.root.after(0, self._redraw_3d_view)
        except Exception:
            pass

    def _redraw_3d_view(self, event=None):
        if not hasattr(self, 'vis_canvas'):
            return
        self.vis_canvas.delete("all")
        w, h = self.vis_canvas.winfo_width(), self.vis_canvas.winfo_height()
        if w < 10 or h < 10:
            return

        if not self.seq_info:
            self.vis_canvas.create_text(w / 2, h / 2,
                                        text="Click 'Estimate' or 'Auto-Calculate'\nto preview the 3D Scan Volume.",
                                        fill=T['text3'], font=('Segoe UI', 10))
            return

        si = self.seq_info
        fov_y, fov_z = si.get('fov_y', 5.0), si.get('fov_z', 5.0)

        x0, x1 = si['x_min'], si['x_max']
        y0, y1 = si['y_min'] - fov_y / 2, si['y_max'] + fov_y / 2
        z0, z1 = si['z_min'] - fov_z / 2, si['z_max'] + fov_z / 2

        dx, dy, dz = max(1e-5, x1 - x0), max(1e-5, y1 - y0), max(1e-5, z1 - z0)
        max_d = max(dx, dy, dz)
        scale = min(w, h) * 0.4 / max_d
        cx_screen, cy_screen = w / 2, h / 2

        angle = math.radians(30)

        def proj(x, y, z):
            nx = x - (x0 + x1) / 2
            ny = y - (y0 + y1) / 2
            nz = z - (z0 + z1) / 2
            u = (ny - nx) * math.cos(angle)
            v = -nz + (nx + ny) * math.sin(angle)
            return cx_screen + u * scale, cy_screen + v * scale

        bb_color = T['accent']

        def draw_cube(x_a, x_b, y_a, y_b, z_a, z_b, fill_top, fill_right, fill_left):
            cx0, cx1 = min(x_a, x_b), max(x_a, x_b)
            cy0, cy1 = min(y_a, y_b), max(y_a, y_b)
            cz0, cz1 = min(z_a, z_b), max(z_a, z_b)
            p000, p100 = proj(cx0, cy0, cz0), proj(cx1, cy0, cz0)
            p010, p110 = proj(cx0, cy1, cz0), proj(cx1, cy1, cz0)
            p001, p101 = proj(cx0, cy0, cz1), proj(cx1, cy0, cz1)
            p011, p111 = proj(cx0, cy1, cz1), proj(cx1, cy1, cz1)

            if fill_left:
                self.vis_canvas.create_polygon(p100, p110, p111, p101, fill=fill_left, outline="")
            if fill_right:
                self.vis_canvas.create_polygon(p010, p110, p111, p011, fill=fill_right, outline="")
            if fill_top:
                self.vis_canvas.create_polygon(p001, p101, p111, p011, fill=fill_top, outline="")

        def get_b(pts, end_idx, fov):
            p1, p2 = pts[0], pts[end_idx]
            return min(p1, p2) - fov / 2, max(p1, p2) + fov / 2

        p000 = proj(x0, y0, z0)
        p100 = proj(x1, y0, z0)
        p010 = proj(x0, y1, z0)
        p110 = proj(x1, y1, z0)
        p001 = proj(x0, y0, z1)
        p101 = proj(x1, y0, z1)
        p011 = proj(x0, y1, z1)
        p111 = proj(x1, y1, z1)

        # — STEP 1: hidden edges (dashed, behind everything) —
        self.vis_canvas.create_line(p100, p110, fill=bb_color, dash=(2, 2))
        self.vis_canvas.create_line(p100, p101, fill=bb_color, dash=(2, 2))
        self.vis_canvas.create_line(p000, p100, fill=bb_color, dash=(2, 2))

        curr_x, curr_y, curr_z = si['cx'], si['cy'], si['cz']
        x_pts, y_pts, z_pts = si['x_pts'], si['y_pts'], si['z_pts']

        try:
            zi = z_pts.index(curr_z)
        except ValueError:
            zi = 0
        try:
            yi = y_pts.index(curr_y)
        except ValueError:
            yi = 0
        try:
            xi = x_pts.index(curr_x)
        except ValueError:
            xi = 0

        # — STEP 2: progression cubes (filled surfaces, no outline) —
        if zi > 0:
            z0_b, z1_b = get_b(z_pts, zi - 1, fov_z)
            draw_cube(min(x_pts), max(x_pts), min(y_pts) - fov_y / 2, max(y_pts) + fov_y / 2, z0_b, z1_b,
                      "#1E3A5F", "#172C46", "#0F1D30")

        if yi > 0:
            y0_b, y1_b = get_b(y_pts, yi - 1, fov_y)
            draw_cube(min(x_pts), max(x_pts), y0_b, y1_b, curr_z - fov_z / 2, curr_z + fov_z / 2,
                      "#1E3A5F", "#172C46", "#0F1D30")

        if xi > 0:
            x0_b, x1_b = min(x_pts[0], curr_x), max(x_pts[0], curr_x)
            draw_cube(x0_b, x1_b, curr_y - fov_y / 2, curr_y + fov_y / 2, curr_z - fov_z / 2, curr_z + fov_z / 2,
                      "#1E3A5F", "#172C46", "#0F1D30")

        # — STEP 3: current FOV frame —
        sr_bl = proj(curr_x, curr_y - fov_y / 2, curr_z - fov_z / 2)
        sr_br = proj(curr_x, curr_y + fov_y / 2, curr_z - fov_z / 2)
        sr_tr = proj(curr_x, curr_y + fov_y / 2, curr_z + fov_z / 2)
        sr_tl = proj(curr_x, curr_y - fov_y / 2, curr_z + fov_z / 2)
        self.vis_canvas.create_polygon(sr_bl, sr_br, sr_tr, sr_tl, fill="#1C3A3A", outline=T['accent'], width=2,
                                       stipple="gray50")

        # — STEP 4: visible edges of the bounding box (on top of everything) —
        self.vis_canvas.create_line(p000, p010, fill=bb_color)
        self.vis_canvas.create_line(p010, p110, fill=bb_color)
        self.vis_canvas.create_line(p000, p001, fill=bb_color)
        self.vis_canvas.create_line(p010, p011, fill=bb_color)
        self.vis_canvas.create_line(p110, p111, fill=bb_color)
        self.vis_canvas.create_line(p001, p101, fill=bb_color)
        self.vis_canvas.create_line(p101, p111, fill=bb_color)
        self.vis_canvas.create_line(p011, p111, fill=bb_color)
        self.vis_canvas.create_line(p001, p011, fill=bb_color)

        # — STEP 5: Start/End labels at the true first/last points of the scan —
        p_start = proj(x_pts[0], y_pts[0], z_pts[0])
        p_end   = proj(x_pts[-1], y_pts[-1], z_pts[-1])
        self.vis_canvas.create_text(p_start[0], p_start[1] + 15, text="Start", fill=T['success'], font=('Segoe UI', 8, 'bold'))
        self.vis_canvas.create_text(p_end[0],   p_end[1]   - 15, text="End",   fill=T['error'],   font=('Segoe UI', 8, 'bold'))

    # --- OTHER WIDGETS ---
    def _create_magnification_widgets(self, parent):
        # Current M position
        coord_frame = ttk.Frame(parent)
        coord_frame.pack(fill=tk.X, pady=2, padx=5)
        ttk.Label(coord_frame, text="M Position:", style='CoordKey.TLabel').pack(side=tk.LEFT)
        self.a_pos_label = ttk.Label(coord_frame, text="0.000", font=("Consolas", 12, "bold"),
                                     foreground=T['accent'], background=T['panel'])
        self.a_pos_label.pack(side=tk.LEFT, padx=5)
        ttk.Label(coord_frame, text="mm", style='CoordKey.TLabel').pack(side=tk.LEFT)
        ttk.Button(coord_frame, text="Reset Zero M", command=self.reset_zero_m).pack(side=tk.RIGHT, padx=5)

        # Total extension (start_tube + M)
        ext_frame = ttk.Frame(parent)
        ext_frame.pack(fill=tk.X, pady=(0, 4), padx=5)
        ttk.Label(ext_frame, text="Total Extension:", style='CoordKey.TLabel').pack(side=tk.LEFT)
        self.ext_label = ttk.Label(ext_frame, text="—", font=("Consolas", 11, "bold"),
                                   foreground=T['text2'], background=T['panel'])
        self.ext_label.pack(side=tk.LEFT, padx=5)
        ttk.Label(ext_frame, text="mm", style='CoordKey.TLabel').pack(side=tk.LEFT)

        # M- / progressbar / M+ buttons
        m_frame = ttk.Frame(parent)
        m_frame.pack(fill=tk.X, pady=6, padx=5)
        self.btn_m_minus = ttk.Button(m_frame, text="◀ M-", command=lambda: self.move_m(-1))
        self.btn_m_minus.pack(side=tk.LEFT, padx=5)
        self.m_progress = ttk.Progressbar(m_frame, orient=tk.HORIZONTAL, mode='determinate', maximum=200)
        self.m_progress.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        self.btn_m_plus = ttk.Button(m_frame, text="M+ ▶", command=lambda: self.move_m(1))
        self.btn_m_plus.pack(side=tk.RIGHT, padx=5)

        # Optical info (calculated magnification)
        self.m_info_label = ttk.Label(parent, text="Optical Engine: Waiting...", font=("Segoe UI", 9, "bold"),
                                      foreground=T['accent'], background=T['panel'])
        self.m_info_label.pack(side=tk.TOP, pady=4)

        # M axis step
        m_step_frame = ttk.Frame(parent)
        m_step_frame.pack(fill=tk.X, padx=5, pady=(0, 6))
        ttk.Label(m_step_frame, text="M Step (mm):", style='CoordKey.TLabel').pack(side=tk.LEFT, padx=(0, 6))
        ttk.Entry(m_step_frame, textvariable=self.step_m_var, width=8).pack(side=tk.LEFT)

    def _create_cnc_control_widgets(self, parent):
        for i in range(4):
            parent.columnconfigure(i, weight=1)

        ttk.Label(parent, text="Serial Port:").grid(row=0, column=0, padx=5, pady=2, sticky="e")
        self.com_port_selector = ttk.Combobox(parent, textvariable=self.com_port_var, state="readonly")
        self.com_port_selector.grid(row=0, column=1, padx=5, pady=2, sticky="ew")
        self.connect_cnc_button = ttk.Button(parent, text="Connect CNC", command=self.toggle_connect_cnc)
        self.connect_cnc_button.grid(row=0, column=2, padx=5, pady=2, sticky="ew")
        self.refresh_com_button = ttk.Button(parent, text="🔄", command=self.update_com_ports)
        self.refresh_com_button.grid(row=0, column=3, padx=2, pady=2, sticky="w")

        ttk.Label(parent, text="Speed (mm/min):").grid(row=1, column=0, padx=5, pady=2, sticky="e")
        ttk.Entry(parent, textvariable=self.speed_var).grid(row=1, column=1, padx=5, pady=2, sticky="ew")
        ttk.Label(parent, text="Step (mm):").grid(row=1, column=2, padx=5, pady=2, sticky="e")
        ttk.Entry(parent, textvariable=self.step_var).grid(row=1, column=3, padx=5, pady=2, sticky="ew")

        ttk.Button(parent, text="Reset Zero XYZ", command=self.reset_zero_xyz).grid(row=2, column=0, columnspan=4,
                                                                                    padx=5, pady=5, sticky="ew")

        move_btn_frame = ttk.Frame(parent)
        move_btn_frame.grid(row=3, column=0, columnspan=4, pady=5, sticky="ew")
        for i in range(5):
            move_btn_frame.columnconfigure(i, weight=1)

        ttk.Button(move_btn_frame, text="Y+", command=lambda: self.manual_move("y", 1)).grid(row=0, column=1, padx=2,
                                                                                             pady=2, sticky="ew")
        ttk.Button(move_btn_frame, text="Z+", command=lambda: self.manual_move("z", 1)).grid(row=0, column=3, padx=2,
                                                                                             pady=2, sticky="ew")
        ttk.Button(move_btn_frame, text="X-", command=lambda: self.manual_move("x", -1)).grid(row=1, column=0, padx=2,
                                                                                              pady=2, sticky="ew")
        ttk.Button(move_btn_frame, text="X+", command=lambda: self.manual_move("x", 1)).grid(row=1, column=2, padx=2,
                                                                                             pady=2, sticky="ew")
        ttk.Button(move_btn_frame, text="Y-", command=lambda: self.manual_move("y", -1)).grid(row=2, column=1, padx=2,
                                                                                              pady=2, sticky="ew")
        ttk.Button(move_btn_frame, text="Z-", command=lambda: self.manual_move("z", -1)).grid(row=2, column=3, padx=2,
                                                                                              pady=2, sticky="ew")

        coord_frame = ttk.Frame(parent)
        coord_frame.grid(row=4, column=0, columnspan=4, pady=5)
        ttk.Label(coord_frame, text="X", style='CoordKey.TLabel').pack(side=tk.LEFT, padx=(2,0))
        self.x_pos_label = ttk.Label(coord_frame, text="0.000", width=7, style='Coord.TLabel')
        self.x_pos_label.pack(side=tk.LEFT, padx=(1, 8))
        ttk.Label(coord_frame, text="Y", style='CoordKey.TLabel').pack(side=tk.LEFT, padx=(2,0))
        self.y_pos_label = ttk.Label(coord_frame, text="0.000", width=7, style='Coord.TLabel')
        self.y_pos_label.pack(side=tk.LEFT, padx=(1, 8))
        ttk.Label(coord_frame, text="Z", style='CoordKey.TLabel').pack(side=tk.LEFT, padx=(2,0))
        self.z_pos_label = ttk.Label(coord_frame, text="0.000", width=7, style='Coord.TLabel')
        self.z_pos_label.pack(side=tk.LEFT, padx=(1, 2))

        self.last_move_label = ttk.Label(parent, text="Last move: N/A", style='Dim.TLabel')
        self.last_move_label.grid(row=5, column=0, columnspan=4, pady=2, sticky="w")

    def _create_sequence_control_widgets(self, parent):
        for i in range(4):
            parent.columnconfigure(i, weight=1)

        ttk.Label(parent, text="Start Position (X,Y,Z):").grid(row=0, column=0, columnspan=3, sticky="w", padx=5,
                                                               pady=2)
        ttk.Entry(parent, textvariable=self.start_pos_vars["x"]).grid(row=1, column=0, padx=2, sticky="ew")
        ttk.Entry(parent, textvariable=self.start_pos_vars["y"]).grid(row=1, column=1, padx=2, sticky="ew")
        ttk.Entry(parent, textvariable=self.start_pos_vars["z"]).grid(row=1, column=2, padx=2, sticky="ew")
        ttk.Button(parent, text="Set Start", command=self.set_start_pos).grid(row=1, column=3, padx=5, sticky="ew")

        ttk.Label(parent, text="End Position (X,Y,Z):").grid(row=2, column=0, columnspan=3, sticky="w", padx=5, pady=2)
        ttk.Entry(parent, textvariable=self.end_pos_vars["x"]).grid(row=3, column=0, padx=2, sticky="ew")
        ttk.Entry(parent, textvariable=self.end_pos_vars["y"]).grid(row=3, column=1, padx=2, sticky="ew")
        ttk.Entry(parent, textvariable=self.end_pos_vars["z"]).grid(row=3, column=2, padx=2, sticky="ew")
        ttk.Button(parent, text="Set End", command=self.set_end_pos).grid(row=3, column=3, padx=5, sticky="ew")

        ttk.Label(parent, text="Step (X,Y,Z) (mm):").grid(row=4, column=0, columnspan=3, sticky="w", padx=5, pady=2)
        ttk.Entry(parent, textvariable=self.step_seq_vars["x"]).grid(row=5, column=0, padx=2, sticky="ew")
        ttk.Entry(parent, textvariable=self.step_seq_vars["y"]).grid(row=5, column=1, padx=2, sticky="ew")
        ttk.Entry(parent, textvariable=self.step_seq_vars["z"]).grid(row=5, column=2, padx=2, sticky="ew")

        ov_frame = ttk.Frame(parent)
        ov_frame.grid(row=6, column=0, columnspan=4, sticky="ew", pady=5)
        ttk.Label(ov_frame, text="Lat. Overlap (%):").pack(side=tk.LEFT, padx=2)
        ttk.Entry(ov_frame, textvariable=self.overlap_lat_var, width=5).pack(side=tk.LEFT, padx=2)
        ttk.Label(ov_frame, text="Focus Overlap (%):").pack(side=tk.LEFT, padx=(10, 2))
        ttk.Entry(ov_frame, textvariable=self.overlap_focus_var, width=5).pack(side=tk.LEFT, padx=2)

        calc_frame = ttk.Frame(parent)
        calc_frame.grid(row=7, column=0, columnspan=4, sticky="ew", pady=5)
        self.auto_calc_button = ttk.Button(calc_frame, text="🪄 Auto-Calculate Steps & Bounds",
                                           command=self.auto_calculate_sequence)
        self.auto_calc_button.pack(fill=tk.X, expand=True, padx=5)

        delay_frame = ttk.Frame(parent)
        delay_frame.grid(row=8, column=0, columnspan=4, sticky="ew", pady=5)
        ttk.Label(delay_frame, text="Delay (s):").pack(side=tk.LEFT, padx=5)
        ttk.Entry(delay_frame, textvariable=self.delay_var, width=7).pack(side=tk.LEFT, padx=2)
        self.estimate_button = ttk.Button(delay_frame, text="Estimate Time & Info", command=self.estimate_sequence_time)
        self.estimate_button.pack(side=tk.LEFT, padx=10)

        btn_seq_frame = ttk.Frame(parent)
        btn_seq_frame.grid(row=9, column=0, columnspan=4, pady=10, sticky="ew")
        self.start_sequence_button = ttk.Button(btn_seq_frame, text="Start Sequence",
                                                command=self.toggle_start_pause_sequence)
        self.start_sequence_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        self.skip_line_button = ttk.Button(btn_seq_frame, text="Skip Y Line", command=self._request_skip_line,
                                           state=tk.DISABLED)
        self.skip_line_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)

        self.progress_bar = ttk.Progressbar(parent, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=10, column=0, columnspan=4, sticky="ew", padx=5, pady=5)
        self.progress_label = ttk.Label(parent, text="Photos: 0/0 | Elapsed: 00h00m00s | ETA: 00h00m00s",
                                        font=("Segoe UI", 9, "bold"))
        self.progress_label.grid(row=11, column=0, columnspan=4, sticky="w", padx=5, pady=2)

        pause_stop_frame = ttk.Frame(parent)
        pause_stop_frame.grid(row=12, column=0, columnspan=4, pady=5, sticky="ew")
        self.pause_sequence_button = ttk.Button(pause_stop_frame, text="Pause", command=self.pause_sequence,
                                                state=tk.DISABLED)
        self.pause_sequence_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        self.resume_sequence_button = ttk.Button(pause_stop_frame, text="Resume", command=self.resume_sequence,
                                                 state=tk.DISABLED)
        self.resume_sequence_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
        self.stop_sequence_button = ttk.Button(pause_stop_frame, text="Stop", command=self.stop_sequence_completely,
                                               state=tk.DISABLED)
        self.stop_sequence_button.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)

    def _create_grbl_console_widgets(self, parent):
        self.grbl_console_output = tk.Text(
            parent, wrap=tk.WORD, state=tk.DISABLED, height=12,
            bg=T['bg'], fg=T['text2'], font=('Consolas', 9),
            insertbackground=T['accent'], selectbackground=T['border'],
            selectforeground=T['text'], relief='flat', borderwidth=0,
            padx=6, pady=4)
        self.grbl_console_output.tag_configure("out", foreground=T['accent'])
        self.grbl_console_output.tag_configure("in",  foreground=T['success'])

        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.grbl_console_output.yview)
        self.grbl_console_output.config(yscrollcommand=scrollbar.set)

        input_frame = ttk.Frame(parent)
        self.grbl_input_var = tk.StringVar()
        self.grbl_input_entry = ttk.Entry(input_frame, textvariable=self.grbl_input_var)
        self.grbl_input_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)
        self.grbl_input_entry.bind("<Return>", self._send_grbl_from_console)

        send_button = ttk.Button(input_frame, text="Send", command=self._send_grbl_from_console)
        send_button.pack(side=tk.LEFT, padx=(5, 0))

        input_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=(2, 5))
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.grbl_console_output.pack(expand=True, fill=tk.BOTH, padx=5, pady=(5, 0))

    def _create_history_widgets(self, parent):
        text_frame = ttk.Frame(parent)
        text_frame.pack(expand=True, fill=tk.BOTH, padx=5, pady=5)

        self.history_text = tk.Text(
            text_frame, wrap=tk.WORD, state=tk.DISABLED, height=8,
            bg=T['bg'], fg=T['text2'], font=('Segoe UI', 9),
            insertbackground=T['accent'], selectbackground=T['border'],
            selectforeground=T['text'], relief='flat', borderwidth=0,
            padx=6, pady=4)
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.history_text.yview)
        self.history_text.config(yscrollcommand=scrollbar.set)

        self.history_text.tag_configure("error",   foreground=T['error'],   font=('Segoe UI', 9, 'bold'))
        self.history_text.tag_configure("warning", foreground=T['warning'])
        self.history_text.tag_configure("info",    foreground=T['info'])

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.history_text.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)

    # --- LOGIC & COMMUNICATION ---
    def trigger_camera_manually(self):
        if not self.connected_cnc:
            messagebox.showwarning("Warning", "CNC not connected.")
            return
        if self.sequence_running:
            messagebox.showwarning("Warning", "Cannot trigger manually while sequence is running.")
            return
        threading.Thread(target=self._manual_trigger_worker, daemon=True).start()

    def _manual_trigger_worker(self):
        self.log_history("Manual Trigger...", level="info")
        self.send_gcode("M64 P0", quiet=False)
        time.sleep(0.1)
        self.send_gcode("M65 P0", quiet=False)

    def _log_to_grbl_console(self, message, direction="in"):
        self.grbl_console_output.config(state=tk.NORMAL)
        self.grbl_console_output.insert(tk.END, message + "\n", direction)
        num_lines = int(self.grbl_console_output.index('end-1c').split('.')[0])
        if num_lines > 500:
            self.grbl_console_output.delete("1.0", f"{num_lines - 500 + 1}.0")
        self.grbl_console_output.see(tk.END)
        self.grbl_console_output.config(state=tk.DISABLED)

    def _send_grbl_from_console(self, event=None):
        command = self.grbl_input_var.get()
        if not command or not self.connected_cnc:
            return
        threading.Thread(target=self._send_grbl_worker, args=(command,), daemon=True).start()
        self.grbl_input_var.set("")

    def _send_grbl_worker(self, command):
        self.send_gcode(command, quiet=False)

    def send_gcode(self, cmd, quiet=False):
        if not (self.connected_cnc and self.cnc and self.cnc.is_open):
            return None
        with self._serial_lock:
            try:
                if cmd != "?":
                    self.root.after(0, self._log_to_grbl_console, f"> {cmd}", "out")
                self.cnc.write((cmd + '\n').encode('utf-8'))
                response_lines = []
                timeout_duration = 0.25 if cmd == "?" else 2.0
                end_time = time.time() + timeout_duration
                while time.time() < end_time:
                    if self.cnc.in_waiting > 0:
                        line = self.cnc.readline().decode('utf-8', errors='ignore').strip()
                        if line:
                            response_lines.append(line)
                            if (cmd == "?" and line.startswith("<") and line.endswith(">")) or \
                                    (cmd != "?" and ('ok' in line.lower() or 'error' in line.lower())):
                                break
                    else:
                        time.sleep(0.005)
                response = "\n".join(response_lines)
                if cmd != "?" and response:
                    self.root.after(0, self._log_to_grbl_console, f"< {response}", "in")
                return response
            except Exception as e:
                self.log_history(f"ERROR: CNC Communication Error: {e}", level="error")
                self.toggle_connect_cnc(interactive=False)
                return None

    def update_com_ports(self, select_last=False):
        ports = [port.device for port in serial.tools.list_ports.comports()]
        self.com_port_selector['values'] = ports
        if ports:
            if select_last and self.com_port_var.get() == "":
                self.com_port_var.set(ports[-1])
            elif not self.com_port_var.get() in ports:
                self.com_port_var.set(ports[0])
        else:
            self.com_port_var.set("")

    def toggle_connect_cnc(self, interactive=True):
        if not self.connected_cnc:
            try:
                port = self.com_port_var.get()
                if not port:
                    if interactive:
                        messagebox.showerror("Error", "No Port selected.")
                    return False
                self.cnc = serial.Serial(port, 115200, timeout=1, write_timeout=2)
                time.sleep(1)
                self.cnc.reset_input_buffer()
                self.cnc.reset_output_buffer()
                self.cnc.write(b"\r\n\r\n")
                time.sleep(0.5)
                if self.cnc.in_waiting > 0:
                    self.cnc.read(self.cnc.in_waiting)
                self.send_gcode("$X", quiet=True)
                self.send_gcode("G90", quiet=True)
                self.connected_cnc = True
                self.connect_cnc_button.config(text="Disconnect CNC")

                self.root.after(100, self.query_position_loop)
                if interactive:
                    messagebox.showinfo("CNC", f"Connected to {port}")
            except Exception as e:
                if interactive:
                    messagebox.showerror("CNC Connection Error", f"Connection failed: {e}")
                if self.cnc:
                    self.cnc.close()
                self.cnc = None
                return False
        else:
            if self.cnc and self.cnc.is_open:
                self.cnc.close()
            self.cnc = None
            self.connected_cnc = False
            self.connect_cnc_button.config(text="Connect CNC")
            self.current_pos = {"x": 0.0, "y": 0.0, "z": 0.0, "a": 0.0}
            self.update_position_display()
            if interactive:
                messagebox.showinfo("CNC", "Disconnected")
        self.update_sequence_buttons_state()
        return self.connected_cnc

    # ==============================================================================
    # ⚡ NON-BLOCKING ASYNC POLLING
    # ==============================================================================
    def query_position_loop(self):
        if getattr(self, 'sequence_running', False):
            self.root.after(200, self.query_position_loop)
            return

        if self.connected_cnc and getattr(self, 'cnc', None) and self.cnc.is_open:
            if not getattr(self, '_is_polling', False):
                self._is_polling = True
                threading.Thread(target=self._async_status_request, daemon=True).start()

        self.root.after(50, self.query_position_loop)

    def _async_status_request(self):
        try:
            status_line = self.send_gcode("?", quiet=True)
            if status_line:
                px, py, pz, pa, parsed_ok = self._parse_grbl_status(status_line)
                if parsed_ok:
                    if (abs(self.current_pos["x"] - px) > 0.001 or
                            abs(self.current_pos["y"] - py) > 0.001 or
                            abs(self.current_pos["z"] - pz) > 0.001 or
                            abs(self.current_pos["a"] - pa) > 0.001):
                        self.current_pos = {"x": px, "y": py, "z": pz, "a": pa}
                        self.root.after(0, self.update_position_display)
        except Exception as e:
            self.log_history(f"WARNING: Serial communication lost ({e}).", level="error")
            self.connected_cnc = False
            self.root.after(0, lambda: self.connect_cnc_button.config(text="Connect CNC"))
        finally:
            self._is_polling = False

    def _parse_grbl_status(self, status_line):
        wpos_x, wpos_y, wpos_z, wpos_a = None, None, None, None
        mpos_x, mpos_y, mpos_z, mpos_a = None, None, None, None
        wco_x, wco_y, wco_z, wco_a = None, None, None, None
        try:
            parts = status_line.strip('<>').split('|')
            for part in parts:
                if part.startswith("WPos:"):
                    coords = list(map(float, part[5:].split(',')))
                    wpos_x, wpos_y, wpos_z = coords[0], coords[1], coords[2]
                    wpos_a = coords[3] if len(coords) > 3 else 0.0
                    return wpos_x, wpos_y, wpos_z, wpos_a, True
                elif part.startswith("MPos:"):
                    coords = list(map(float, part[5:].split(',')))
                    mpos_x, mpos_y, mpos_z = coords[0], coords[1], coords[2]
                    mpos_a = coords[3] if len(coords) > 3 else 0.0
                elif part.startswith("WCO:"):
                    coords = list(map(float, part[4:].split(',')))
                    wco_x, wco_y, wco_z = coords[0], coords[1], coords[2]
                    wco_a = coords[3] if len(coords) > 3 else 0.0
            if all(v is not None for v in [mpos_x, mpos_y, mpos_z, wco_x, wco_y, wco_z]):
                wco_a = wco_a if wco_a is not None else 0.0
                mpos_a = mpos_a if mpos_a is not None else 0.0
                return mpos_x - wco_x, mpos_y - wco_y, mpos_z - wco_z, mpos_a - wco_a, True
        except Exception:
            return None, None, None, None, False
        return None, None, None, None, False

    def update_position_display(self):
        self.x_pos_label.config(text=f'{self.current_pos["x"]:.3f}')
        self.y_pos_label.config(text=f'{self.current_pos["y"]:.3f}')
        self.z_pos_label.config(text=f'{self.current_pos["z"]:.3f}')
        self.a_pos_label.config(text=f'{self.current_pos["a"]:.3f}')

        m_max = 200.0
        prof_m = self.machine_profiles.get(self.active_machine_var.get())
        if prof_m:
            try:
                m_max = float(prof_m.get("travel_m", 200.0))
            except Exception:
                pass

        self.m_progress.config(maximum=m_max)
        self.m_progress['value'] = max(0.0, min(m_max, self.current_pos["a"]))

        prof = self.lens_profiles.get(self.active_lens_var.get())
        if prof:
            try:
                nom_mag    = float(prof.get("nominal_mag", 10.0))
                obj_tube   = float(prof.get("obj_tube_length", 160.0))
                start_tube = float(prof.get("start_tube_length", 160.0))

                total_ext = start_tube + self.current_pos["a"]
                self.ext_label.config(text=f"{total_ext:.3f}")

                f = OpticalCalculator.get_focal_length(nom_mag, obj_tube)
                current_i = total_ext

                if current_i > f:
                    current_mag = (current_i - f) / f
                    self.m_info_label.config(text=f"Current Magnification: {current_mag:.2f}x")
                else:
                    self.m_info_label.config(text="Optical Error: Extension too short")
            except Exception:
                self.ext_label.config(text="—")
                self.m_info_label.config(text="Optical Engine: Profile Error")
        else:
            self.ext_label.config(text="—")
            self.m_info_label.config(text="Optical Engine: No profile selected")

    def reset_zero_xyz(self):
        response = self.send_gcode("G92 X0 Y0 Z0", quiet=False)
        if response and 'ok' in response.lower():
            self.current_pos["x"] = 0.0
            self.current_pos["y"] = 0.0
            self.current_pos["z"] = 0.0
            self.update_position_display()
            self.last_move_label.config(text="Zero Reset XYZ")
        else:
            self.last_move_label.config(text="Zero Reset Failed")

    def reset_zero_m(self):
        response = self.send_gcode("G92 A0", quiet=False)
        if response and 'ok' in response.lower():
            self.current_pos["a"] = 0.0
            self.update_position_display()
            self.last_move_label.config(text="Zero Reset M")
        else:
            self.last_move_label.config(text="Zero Reset Failed")

    def move_m(self, direction):
        if not self.connected_cnc:
            return
        threading.Thread(target=self._move_m_worker, args=(direction,), daemon=True).start()

    def _move_m_worker(self, direction):
        try:
            step = float(self.step_m_var.get())
        except ValueError:
            self.root.after(0, lambda: messagebox.showerror("Input Error", "Invalid step size."))
            return

        m_current = self.current_pos["a"]
        m_target = m_current + (direction * step)
        prof = self.lens_profiles.get(self.active_lens_var.get())

        if prof:
            nom_mag = float(prof.get("nominal_mag", 10.0))
            obj_tube = float(prof.get("obj_tube_length", 160.0))
            start_tube = float(prof.get("start_tube_length", 160.0))
            invert = prof.get("invert_x", False)

            f = OpticalCalculator.get_focal_length(nom_mag, obj_tube)
            o_current = OpticalCalculator.calculate_working_distance(m_current, f, start_tube)
            o_target = OpticalCalculator.calculate_working_distance(m_target, f, start_tube)

            if o_current is not None and o_target is not None:
                delta_o = o_target - o_current
                dx = -delta_o if invert else delta_o
                self.send_gcode("G91", quiet=True)
                self.send_gcode(f"G0 A{direction * step:.3f} X{dx:.4f}", quiet=False)
                self.send_gcode("G90", quiet=True)
                msg = f"M move: A{(direction * step):+.2f}mm (Optical Comp X: {dx:+.4f}mm)"
                self.root.after(0, lambda m=msg: self.last_move_label.config(text=m))
                return

        self.send_gcode("G91", quiet=True)
        self.send_gcode(f"G0 A{direction * step:.3f}", quiet=False)
        self.send_gcode("G90", quiet=True)
        msg = f"M move: A{(direction * step):+.2f}mm (No Comp)"
        self.root.after(0, lambda m=msg: self.last_move_label.config(text=m))

    def manual_move(self, axis, direction):
        if not self.connected_cnc:
            return
        threading.Thread(target=self._manual_move_worker, args=(axis, direction), daemon=True).start()

    def _manual_move_worker(self, axis, direction):
        try:
            step = float(self.step_var.get())
        except ValueError:
            self.root.after(0, lambda: messagebox.showerror("Input Error", "Invalid speed or step."))
            return
        self.send_gcode("G91", quiet=True)
        self.send_gcode(f"G0 {axis.upper()}{direction * step}", quiet=False)
        self.send_gcode("G90", quiet=True)
        op = "+" if direction > 0 else "-"
        self.root.after(0, lambda: self.last_move_label.config(text=f"Move: {axis.upper()}{op}{step}mm"))

    # --- OPTICAL ENGINE CALCULATIONS ---
    def auto_calculate_sequence(self):
        prof_l = self.lens_profiles.get(self.active_lens_var.get())
        prof_c = self.camera_profiles.get(self.active_camera_var.get())
        prof_m = self.machine_profiles.get(self.active_machine_var.get())

        if not (prof_l and prof_c and prof_m):
            messagebox.showerror("Profile Error", "Please select valid Lens, Camera, and Machine profiles.")
            return

        try:
            calc = OpticalCalculator.calculate_steps(
                nom_mag=float(prof_l.get("nominal_mag", 10.0)),
                na=float(prof_l.get("na", 0.25)),
                obj_tube=float(prof_l.get("obj_tube_length", 160.0)),
                start_tube=float(prof_l.get("start_tube_length", 160.0)),
                current_m_pos=self.current_pos["a"],
                cam_w=float(prof_c.get("width", 36.0)),
                cam_h=float(prof_c.get("height", 24.0)),
                coc=float(prof_c.get("coc", 0.03)),
                prec_x=float(prof_m.get("prec_x", 0.01)),
                prec_y=float(prof_m.get("prec_y", 0.01)),
                prec_z=float(prof_m.get("prec_z", 0.001)),
                ov_lat=float(self.overlap_lat_var.get()) / 100.0,
                ov_foc=float(self.overlap_focus_var.get()) / 100.0
            )

            step_x, step_y, step_z = calc["step_x"], calc["step_y"], calc["step_z"]
            self.step_seq_vars["x"].set(f"{step_x:.6g}")
            self.step_seq_vars["y"].set(f"{step_y:.6g}")
            self.step_seq_vars["z"].set(f"{step_z:.6g}")

            for axis, step_val in zip(["x", "y", "z"], [step_x, step_y, step_z]):
                s_val = float(self.start_pos_vars[axis].get())
                e_val = float(self.end_pos_vars[axis].get())
                delta = e_val - s_val

                if abs(delta) > 1e-5:
                    n_intervals = math.ceil(abs(delta) / step_val)
                    new_delta = n_intervals * step_val
                    center = (s_val + e_val) / 2.0
                    if e_val > s_val:
                        new_s = center - (new_delta / 2.0)
                        new_e = center + (new_delta / 2.0)
                    else:
                        new_s = center + (new_delta / 2.0)
                        new_e = center - (new_delta / 2.0)
                    self.start_pos_vars[axis].set(f"{new_s:.6g}")
                    self.end_pos_vars[axis].set(f"{new_e:.6g}")

            self.update_3d_preview_data()
            messagebox.showinfo("Optical Engine",
                                f"Calculations Successful!\n\nMagnification: {calc['M']:.2f}x\nDoF: {calc['dof'] * 1000:.1f} µm\n\nSteps Applied:\nX: {step_x} mm\nY: {step_y} mm\nZ: {step_z} mm\n\nBounds symmetrically expanded to fit exactly.")

        except ValueError as e:
            messagebox.showerror("Math/Input Error", str(e))

    def set_start_pos(self):
        for axis, var in self.start_pos_vars.items():
            var.set(f'{self.current_pos[axis]:.3f}')
        self.update_3d_preview_data()

    def set_end_pos(self):
        for axis, var in self.end_pos_vars.items():
            var.set(f'{self.current_pos[axis]:.3f}')
        self.update_3d_preview_data()

    # ==============================================================================
    # ⏱️ ROBUST ESTIMATION AND PREVIEW
    # ==============================================================================
    def estimate_sequence_time(self):
        try:
            start_pos = {axis: float(var.get()) for axis, var in self.start_pos_vars.items()}
            end_pos = {axis: float(var.get()) for axis, var in self.end_pos_vars.items()}
            steps = {axis: float(var.get()) for axis, var in self.step_seq_vars.items()}
            delay_s = float(self.delay_var.get())

            if steps['x'] <= 0 or steps['y'] <= 0 or steps['z'] <= 0:
                messagebox.showwarning("Warning", "Invalid Steps. Use 'Auto-Calculate' first.")
                return

            prof_m = self.machine_profiles.get(self.active_machine_var.get(), {})
            prec_x = float(prof_m.get("prec_x", 0.01))
            prec_y = float(prof_m.get("prec_y", 0.01))
            prec_z = float(prof_m.get("prec_z", 0.001))

            z_pts = self._generate_scan_points(start_pos['z'], end_pos['z'], steps['z'], prec_z)
            x_pts = self._generate_scan_points(start_pos['x'], end_pos['x'], steps['x'], prec_x)
            y_pts = self._generate_scan_points(start_pos['y'], end_pos['y'], steps['y'], prec_y)

            total_photos = len(z_pts) * len(x_pts) * len(y_pts)

            if total_photos == 0:
                messagebox.showinfo("Estimation", "0 photos computed. Check your Start/End values.")
                return

            time_per_photo = 1.0 + 0.5 + delay_s
            total_seconds = time_per_photo * total_photos
            eta_formatted = time.strftime('%Hh %Mm %Ss', time.gmtime(total_seconds))

            opt_info = ""
            prof_l = self.lens_profiles.get(self.active_lens_var.get())
            prof_c = self.camera_profiles.get(self.active_camera_var.get())

            if prof_l and prof_c:
                try:
                    calc = OpticalCalculator.calculate_steps(
                        nom_mag=float(prof_l.get("nominal_mag", 10.0)),
                        na=float(prof_l.get("na", 0.25)),
                        obj_tube=float(prof_l.get("obj_tube_length", 160.0)),
                        start_tube=float(prof_l.get("start_tube_length", 160.0)),
                        current_m_pos=self.current_pos["a"],
                        cam_w=float(prof_c.get("width", 36.0)),
                        cam_h=float(prof_c.get("height", 24.0)),
                        coc=float(prof_c.get("coc", 0.03)),
                        prec_x=prec_x, prec_y=prec_y, prec_z=prec_z,
                        ov_lat=float(self.overlap_lat_var.get()) / 100.0,
                        ov_foc=float(self.overlap_focus_var.get()) / 100.0
                    )
                    opt_info = f"\n\nCurrent Magnification: {calc['M']:.2f}x\nSensor Footprint (FoV): {calc['fov_y']:.2f} x {calc['fov_z']:.2f} mm\nDepth of Field: {calc['dof'] * 1000:.1f} µm"
                except Exception:
                    pass

            self.progress_label.config(text=f"Photos: 0/{total_photos} | Elapsed: 00h00m00s | ETA: {eta_formatted}")
            self.progress_var.set(0.0)

            self.update_3d_preview_data(x_pts, y_pts, z_pts)

            messagebox.showinfo("Sequence Estimation",
                                f"Total Photos: {total_photos}\nEstimated Duration: ~{eta_formatted}{opt_info}")

        except ValueError:
            messagebox.showerror("Input Error", "Please verify that all input values are valid numbers.")

    def update_sequence_buttons_state(self):
        can_start = self.connected_cnc
        if self.sequence_running:
            paused = self.sequence_is_paused_by_error or self.user_requested_pause_event.is_set()
            self.start_sequence_button.config(state=tk.DISABLED)
            self.skip_line_button.config(state=tk.NORMAL)
            self.pause_sequence_button.config(state=tk.NORMAL if not paused else tk.DISABLED)
            self.resume_sequence_button.config(state=tk.NORMAL if paused else tk.DISABLED)
            self.stop_sequence_button.config(state=tk.NORMAL)
            self.manual_trigger_button.config(state=tk.DISABLED)
            self.btn_m_plus.config(state=tk.DISABLED)
            self.btn_m_minus.config(state=tk.DISABLED)
            self.auto_calc_button.config(state=tk.DISABLED)
        else:
            self.start_sequence_button.config(state=tk.NORMAL if can_start else tk.DISABLED)
            self.skip_line_button.config(state=tk.DISABLED)
            self.pause_sequence_button.config(state=tk.DISABLED)
            self.resume_sequence_button.config(state=tk.DISABLED)
            self.stop_sequence_button.config(state=tk.DISABLED)
            self.manual_trigger_button.config(state=tk.NORMAL if can_start else tk.DISABLED)
            self.btn_m_plus.config(state=tk.NORMAL if can_start else tk.DISABLED)
            self.btn_m_minus.config(state=tk.NORMAL if can_start else tk.DISABLED)
            self.auto_calc_button.config(state=tk.NORMAL if can_start else tk.DISABLED)

    def _handle_sequence_pause(self, message, can_retry_operation=True):
        self.log_history(f"SEQUENCE PAUSED: {message}", level="info")
        self.sequence_is_paused_by_error = not self.user_requested_pause_event.is_set()
        self.root.after(0, self.update_sequence_buttons_state)

        if self.sequence_is_paused_by_error:
            self.root.after(0, lambda m=message: messagebox.showwarning("Sequence Paused", m))

        self.sequence_pause_event.clear()
        self.sequence_pause_event.wait()

        self.sequence_is_paused_by_error = False
        was_manual_pause = self.user_requested_pause_event.is_set()
        self.user_requested_pause_event.clear()
        self.root.after(0, self.update_sequence_buttons_state)

        if self.stop_sequence_flag.is_set():
            return False

        if self.sequence_resume_event.is_set():
            self.sequence_resume_event.clear()
            self.log_history("Resuming sequence.")
            return can_retry_operation and not was_manual_pause

        return False

    def _generate_scan_points(self, start, end, step, precision=None):
        if abs(step) < 1e-5:
            return [start] if start == end else [start, end]
        actual_step = abs(step)

        if precision is not None and precision > 0:
            if actual_step <= 20 * precision + 1e-9:
                quantized = math.floor(actual_step / precision) * precision
                if quantized < precision:
                    quantized = precision
                actual_step = quantized

        points = []
        current = start
        direction = 1 if end >= start else -1
        step_with_direction = actual_step * direction
        count = 0
        epsilon = 1e-7

        while direction * current <= direction * end + epsilon and count < 10000:
            points.append(round(current, 4))
            current += step_with_direction
            count += 1

        if not points or abs(points[-1] - end) > 1e-5:
            points.append(round(end, 4))

        return points

    # ==============================================================================
    # 🔒 SECURE WORKER WITH CSV LOGGING
    # ==============================================================================
    def _sequence_worker(self, seq_config):
        self.log_history("Sequence started.", level="info")

        start_pos = seq_config['start_pos']
        end_pos = seq_config['end_pos']
        steps = seq_config['steps']
        full_delay_s = seq_config['delay']
        sequence_speed = seq_config['speed']
        precisions = seq_config['precisions']

        half_delay = full_delay_s / 2.0

        z_points = self._generate_scan_points(start_pos['z'], end_pos['z'], steps['z'], precisions['z'])
        x_points = self._generate_scan_points(start_pos['x'], end_pos['x'], steps['x'], precisions['x'])
        y_points = self._generate_scan_points(start_pos['y'], end_pos['y'], steps['y'], precisions['y'])

        # Warning if an axis has reached the 10 000 point limit (step size is likely too small)
        for axis_name, pts, start_v, end_v, step_v in [
            ('Z', z_points, start_pos['z'], end_pos['z'], steps['z']),
            ('X', x_points, start_pos['x'], end_pos['x'], steps['x']),
            ('Y', y_points, start_pos['y'], end_pos['y'], steps['y']),
        ]:
            span = abs(end_v - start_v)
            if step_v > 1e-5 and span / step_v > 9999:
                self.log_history(
                    f"WARNING Axis {axis_name}: step size ({step_v:.4f} mm) is very small for "
                    f"the span ({span:.3f} mm) — sequence limited to 10 000 points. "
                    "Check your parameters.", level="warning"
                )

        total_photos = len(z_points) * len(x_points) * len(y_points)
        photos_taken = 0
        start_time = time.time()

        # Initialize CSV Logging module
        log_filepath = None
        log_file     = None
        log_writer   = None

        # Calculate actual magnification at launch time (M-axis fixed during the sequence)
        actual_magnification = None
        try:
            prof_l_calc  = self.lens_profiles.get(self.active_lens_var.get(), {})
            nom_mag_c    = float(prof_l_calc.get("nominal_mag", 10.0))
            obj_tube_c   = float(prof_l_calc.get("obj_tube_length", 160.0))
            start_tube_c = float(prof_l_calc.get("start_tube_length", 160.0))
            f_c = OpticalCalculator.get_focal_length(nom_mag_c, obj_tube_c)
            i_c = start_tube_c + self.current_pos["a"]
            if i_c > f_c:
                actual_magnification = round((i_c - f_c) / f_c, 4)
        except Exception:
            pass

        if self.log_enabled_var.get():
            try:
                log_dir = self.log_dir_var.get()
                if not os.path.exists(log_dir):
                    os.makedirs(log_dir)

                timestamp_str = time.strftime("%Y%m%d_%H%M%S")
                log_filepath  = os.path.join(log_dir, f"session_log_{timestamp_str}.csv")

                log_file   = open(log_filepath, 'w', newline='', encoding='utf-8')
                log_writer = csv.writer(log_file)

                prof_l = self.lens_profiles.get(self.active_lens_var.get(), {})
                prof_c = self.camera_profiles.get(self.active_camera_var.get(), {})
                mag_str = f"{actual_magnification}x" if actual_magnification is not None else "Unknown"

                log_writer.writerow(["# METADATA: Session Started at", time.strftime("%Y-%m-%d %H:%M:%S")])
                total_extension = float(prof_l.get("start_tube_length", 160.0)) + self.current_pos.get("a", 0.0)
                log_writer.writerow(["# LENS:", prof_l.get("name", "Unknown"),
                                     f"NA={prof_l.get('na', 'Unknown')}",
                                     f"NomMag={prof_l.get('nominal_mag', 'Unknown')}",
                                     f"ActualMag={mag_str}",
                                     f"Extension={total_extension:.2f}mm"])
                log_writer.writerow(["# CAMERA:", prof_c.get("name", "Unknown"),
                                     f"Sensor={prof_c.get('width', 0)}x{prof_c.get('height', 0)}mm"])
                log_writer.writerow(["# PARAMS:", f"LatOverlap={self.overlap_lat_var.get()}%",
                                     f"FocOverlap={self.overlap_focus_var.get()}%",
                                     f"Delay={full_delay_s}s"])
                log_writer.writerow(["# TOTAL CALCULATED PHOTOS:", total_photos])
                log_writer.writerow([])
                log_writer.writerow(["frame_index", "timestamp", "x_um", "y_um", "z_um", "magnification"])
                log_file.flush()

                self.log_history(f"SD-Card Log initialized at {log_filepath}")
            except Exception as e:
                self.log_history(f"Failed to initialize CSV log: {e}", "error")
                if log_file:
                    log_file.close()
                log_file   = None
                log_writer = None

        self.seq_info = {
            'x_pts': x_points, 'y_pts': y_points, 'z_pts': z_points,
            'x_min': min(x_points), 'x_max': max(x_points),
            'y_min': min(y_points), 'y_max': max(y_points),
            'z_min': min(z_points), 'z_max': max(z_points),
            'cx': start_pos['x'], 'cy': start_pos['y'], 'cz': start_pos['z'],
            'fov_y': seq_config['fov_y'], 'fov_z': seq_config['fov_z'],
            'start_x': x_points[0] if x_points else start_pos['x']
        }
        self.root.after(0, self._redraw_3d_view)

        self.log_history(f"Calculation: {total_photos} photos planned.")
        self.send_gcode("G90", quiet=True)

        safe_start_x = x_points[0] if x_points else start_pos['x']
        safe_start_y = y_points[0] if y_points else start_pos['y']
        safe_start_z = z_points[0] if z_points else start_pos['z']

        if self.send_gcode(f"G1 X{safe_start_x:.3f} Y{safe_start_y:.3f} Z{safe_start_z:.3f} F{sequence_speed}") is None:
            self._sequence_cleanup()
            return

        if not self._wait_for_move_completion(safe_start_x, safe_start_y, safe_start_z):
            if not self.stop_sequence_flag.is_set():
                self.stop_sequence_flag.set()
            self._sequence_cleanup()
            return

        try:
            for actual_z in z_points:
                if self.stop_sequence_flag.is_set(): break
                for actual_y in y_points:
                    if self.stop_sequence_flag.is_set() or self.skip_line_event.is_set(): break
                    for actual_x in x_points:
                        if self.stop_sequence_flag.is_set() or self.skip_line_event.is_set(): break
                        self._check_manual_pause_request()
                        if self.stop_sequence_flag.is_set() or self.skip_line_event.is_set(): break

                        if half_delay > 0:
                            time.sleep(half_delay)
                        self.send_gcode(f"G1 X{actual_x:.3f} Y{actual_y:.3f} Z{actual_z:.3f} F{sequence_speed}",
                                        quiet=True)
                        if not self._wait_for_move_completion(actual_x, actual_y, actual_z):
                            self.log_history("Move Error", level="error")
                        else:
                            if half_delay > 0:
                                time.sleep(half_delay)

                            # CAMERA TRIGGER
                            self.send_gcode("M64 P0", quiet=False)
                            time.sleep(0.1)
                            self.send_gcode("M65 P0", quiet=False)

                            # CSV LOG WRITING IMMEDIATELY AFTER TRIGGER
                            if log_writer:
                                try:
                                    log_writer.writerow([photos_taken, time.time(),
                                                         round(actual_x * 1000, 2),
                                                         round(actual_y * 1000, 2),
                                                         round(actual_z * 1000, 2),
                                                         actual_magnification if actual_magnification is not None else ""])
                                    log_file.flush()
                                except Exception as e:
                                    self.log_history(f"Log write error: {e}", "warning")

                            self.seq_info['cx'] = actual_x
                            self.seq_info['cy'] = actual_y
                            self.seq_info['cz'] = actual_z
                            self.root.after(0, self._redraw_3d_view)

                        photos_taken += 1
                        progress_percent = (photos_taken / total_photos) * 100 if total_photos > 0 else 0
                        elapsed = time.time() - start_time
                        elapsed_str = time.strftime('%Hh%Mm%Ss', time.gmtime(elapsed))

                        if photos_taken > 0:
                            eta_sec = (total_photos - photos_taken) * (elapsed / photos_taken)
                            eta_str = time.strftime('%Hh%Mm%Ss', time.gmtime(eta_sec))
                        else:
                            eta_str = "Calculating..."

                        prog_txt = f"Photos: {photos_taken}/{total_photos} | Elapsed: {elapsed_str} | ETA: {eta_str} | {progress_percent:.1f}%"
                        self.root.after(0, lambda p=progress_percent, txt=prog_txt: [self.progress_var.set(p),
                                                                                     self.progress_label.config(
                                                                                         text=txt)])

                if self.skip_line_event.is_set():
                    self.log_history("Skipping to next Y-line...", level="info")
                    self.skip_line_event.clear()
        except Exception as e:
            self.log_history(f"CRITICAL SEQ ERROR: {e}", level="error")

        # End of sequence: summary + close CSV file
        if log_file:
            try:
                log_writer.writerow([])
                log_writer.writerow(["# SUMMARY: Sequence Terminated at", time.strftime("%Y-%m-%d %H:%M:%S")])
                log_writer.writerow(["# ACTUAL PHOTOS TAKEN:", photos_taken])
                log_writer.writerow(["# TOTAL DURATION (Seconds):", round(time.time() - start_time, 2)])
                log_file.flush()
            except Exception:
                pass
            finally:
                log_file.close()

        self._sequence_cleanup()
        if not self.stop_sequence_flag.is_set():
            self.log_history("Sequence finished.", level="info")
            play_sound()
        else:
            self.log_history("Sequence stopped.", level="info")

    def _wait_for_move_completion(self, x_coord=None, y_coord=None, z_coord=None, tolerance=0.05, timeout_s=45):
        start_time = time.time()
        while time.time() - start_time < timeout_s:
            if self.stop_sequence_flag.is_set():
                return False
            try:
                status_line = self.send_gcode("?", quiet=True)
                if status_line:
                    # Immediate detection of a GRBL alarm
                    if "alarm" in status_line.lower():
                        self.log_history(
                            f"GRBL ALARM detected: {status_line.strip()} — sequence interrupted. "
                            "Use $X in the GRBL console to unlock the machine.",
                            level="error"
                        )
                        self.stop_sequence_flag.set()
                        return False

                    is_idle = "Idle" in status_line
                    px, py, pz, pa, parsed_ok = self._parse_grbl_status(status_line)
                    if parsed_ok:
                        self.current_pos = {"x": px, "y": py, "z": pz, "a": pa}
                        self.root.after(0, self.update_position_display)

                        if is_idle:
                            if (x_coord is None or abs(px - x_coord) <= tolerance) and \
                                    (y_coord is None or abs(py - y_coord) <= tolerance) and \
                                    (z_coord is None or abs(pz - z_coord) <= tolerance):
                                return True
            except Exception:
                pass
            time.sleep(0.05)
        return False

    def _check_manual_pause_request(self):
        if self.user_requested_pause_event.is_set():
            self._handle_sequence_pause("Manually paused.")

    def _sequence_cleanup(self):
        self.sequence_running = False
        self.sequence_is_paused_by_error = False
        self.user_requested_pause_event.clear()
        self.sequence_pause_event.set()
        self.stop_sequence_flag.clear()
        self.skip_line_event.clear()
        self.root.after(0, self.update_sequence_buttons_state)

    def toggle_start_pause_sequence(self):
        if not self.sequence_running:
            try:
                prof_m = self.machine_profiles.get(self.active_machine_var.get(), {})
                prof_l = self.lens_profiles.get(self.active_lens_var.get(), {})
                prof_c = self.camera_profiles.get(self.active_camera_var.get(), {})

                fov_y, fov_z = 5.0, 5.0
                try:
                    f = OpticalCalculator.get_focal_length(float(prof_l.get("nominal_mag", 10.0)),
                                                           float(prof_l.get("obj_tube_length", 160.0)))
                    current_i = float(prof_l.get("start_tube_length", 160.0)) + self.current_pos["a"]
                    if current_i > f:
                        M = (current_i - f) / f
                        fov_y = float(prof_c.get("width", 36.0)) / M
                        fov_z = float(prof_c.get("height", 24.0)) / M
                except Exception:
                    pass

                seq_config = {
                    'start_pos': {axis: float(var.get()) for axis, var in self.start_pos_vars.items()},
                    'end_pos': {axis: float(var.get()) for axis, var in self.end_pos_vars.items()},
                    'steps': {axis: float(var.get()) for axis, var in self.step_seq_vars.items()},
                    'delay': float(self.delay_var.get()),
                    'speed': self.speed_var.get(),
                    'fov_y': fov_y,
                    'fov_z': fov_z,
                    'precisions': {
                        'x': float(prof_m.get("prec_x", 0.01)),
                        'y': float(prof_m.get("prec_y", 0.01)),
                        'z': float(prof_m.get("prec_z", 0.001))
                    }
                }
            except ValueError:
                messagebox.showerror("Input Error", "Invalid sequence parameters.")
                return

            self.history_text.config(state=tk.NORMAL)
            self.history_text.delete('1.0', tk.END)
            self.history_text.config(state=tk.DISABLED)

            self.sequence_running = True
            self.stop_sequence_flag.clear()
            self.skip_line_event.clear()

            self.sequence_thread = threading.Thread(target=self._sequence_worker, args=(seq_config,), daemon=True)
            self.sequence_thread.start()

        self.update_sequence_buttons_state()

    def pause_sequence(self):
        if self.sequence_running and not (self.sequence_is_paused_by_error or self.user_requested_pause_event.is_set()):
            self.user_requested_pause_event.set()

    def _request_skip_line(self):
        if self.sequence_running:
            self.skip_line_event.set()

    def resume_sequence(self):
        if self.sequence_running and (self.sequence_is_paused_by_error or self.user_requested_pause_event.is_set()):
            self.sequence_resume_event.set()
            self.sequence_pause_event.set()

    def stop_sequence_completely(self):
        if self.sequence_running:
            self.stop_sequence_flag.set()
            if self.sequence_is_paused_by_error or self.user_requested_pause_event.is_set():
                self.sequence_pause_event.set()

    def log_history(self, message, level="normal"):
        self.root.after(0, self._update_history_text, message, level)

    def _update_history_text(self, message, level):
        MAX_HISTORY_LINES = 500
        timestamp = time.strftime("%H:%M:%S")
        full_message = f"[{timestamp}] {message}\n"

        self.history_text.config(state=tk.NORMAL)
        tag_to_use = None
        if "ERROR" in message.upper() or level.lower() == "error":
            tag_to_use = "error"
        elif "WARNING" in message.upper() or level.lower() == "warning":
            tag_to_use = "warning"
        elif level.lower() == "info":
            tag_to_use = "info"

        self.history_text.insert(tk.END, full_message, tag_to_use)
        num_lines = int(self.history_text.index('end-1c').split('.')[0])
        scan_index = 1

        while num_lines > MAX_HISTORY_LINES:
            tags = self.history_text.tag_names(f"{scan_index}.0")
            if "error" in tags:
                scan_index += 1
                if scan_index >= num_lines:
                    break
            else:
                self.history_text.delete(f"{scan_index}.0", f"{scan_index + 1}.0")
                num_lines -= 1

        self.history_text.see(tk.END)
        self.history_text.config(state=tk.DISABLED)

    def on_closing(self):
        try:
            self._save_config()
        except Exception:
            pass

        self.stop_sequence_completely()
        if hasattr(self, 'sequence_thread') and self.sequence_thread.is_alive():
            self.sequence_thread.join(timeout=1.0)

        if self.cnc and self.cnc.is_open:
            self.cnc.close()

        self.root.destroy()

    def _save_config(self):
        sashes = {'main': [], 'left': [], 'center': [], 'right': []}
        try:
            sashes['main'] = [self.main_pw.sashpos(i) for i in range(len(self.main_pw.panes()) - 1)]
            sashes['left'] = [self.left_pw.sashpos(i) for i in range(len(self.left_pw.panes()) - 1)]
            sashes['center'] = [self.center_pw.sashpos(i) for i in range(len(self.center_pw.panes()) - 1)]
            sashes['right'] = [self.right_pw.sashpos(i) for i in range(len(self.right_pw.panes()) - 1)]
        except Exception:
            pass

        config_data = {
            'com_port': self.com_port_var.get(),
            'start_pos': {axis: self.start_pos_vars[axis].get() for axis in "xyz"},
            'end_pos': {axis: self.end_pos_vars[axis].get() for axis in "xyz"},
            'step_seq': {axis: self.step_seq_vars[axis].get() for axis in "xyz"},
            'delay': self.delay_var.get(),
            'cnc_speed': self.speed_var.get(),
            'cnc_step': self.step_var.get(),
            'step_m': self.step_m_var.get(),
            'overlap_lat': self.overlap_lat_var.get(),
            'overlap_focus': self.overlap_focus_var.get(),
            'lens_profiles': self.lens_profiles,
            'active_lens': self.active_lens_var.get(),
            'camera_profiles': self.camera_profiles,
            'active_camera': self.active_camera_var.get(),
            'machine_profiles': self.machine_profiles,
            'active_machine': self.active_machine_var.get(),
            'log_enabled': self.log_enabled_var.get(),
            'log_dir': self.log_dir_var.get(),
            'layout': self.layout,
            'sashes': sashes
        }

        try:
            with open(self.config_file, 'w') as f:
                json.dump(config_data, f, indent=4)
        except Exception as e:
            messagebox.showerror("Save Error",
                                 f"Cannot write configuration file!\n\nPath: {self.config_file}\n\nWindows Error: {e}")

    def _load_config(self):
        defaults = {
            'com_port': '',
            'start_pos': {'x': '0.0', 'y': '0.0', 'z': '0.0'},
            'end_pos': {'x': '10.0', 'y': '10.0', 'z': '0.0'},
            'step_seq': {'x': '1.0', 'y': '1.0', 'z': '1.0'},
            'delay': '0.5',
            'cnc_speed': '1000',
            'cnc_step': '1',
            'step_m': '1',
            'overlap_lat': '40',
            'overlap_focus': '20',
            'lens_profiles': {},
            'active_lens': '',
            'camera_profiles': {},
            'active_camera': '',
            'machine_profiles': {},
            'active_machine': '',
            'log_enabled': True,
            'log_dir': os.path.join(os.path.expanduser("~"), "Desktop", "CNC_Logs"),
            'layout': {"left": ["cnc", "trigger", "log_panel", "mag"], "center": ["hw", "seq", "vis"],
                       "right": ["grbl", "hist"]},
            'sashes': {}
        }

        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    config_data = json.load(f)
            else:
                config_data = defaults
        except Exception:
            config_data = defaults

        self.com_port_var.set(config_data.get('com_port', defaults['com_port']))

        start_pos = config_data.get('start_pos', defaults['start_pos'])
        for axis in "xyz":
            self.start_pos_vars[axis].set(start_pos.get(axis, "0.0"))

        end_pos = config_data.get('end_pos', defaults['end_pos'])
        step_seq = config_data.get('step_seq', defaults['step_seq'])
        for axis in "xyz":
            self.end_pos_vars[axis].set(end_pos.get(axis, "0.0"))
            self.step_seq_vars[axis].set(step_seq.get(axis, "1.0"))

        self.delay_var.set(config_data.get('delay', defaults['delay']))
        self.speed_var.set(config_data.get('cnc_speed', defaults['cnc_speed']))
        self.step_var.set(config_data.get('cnc_step', defaults['cnc_step']))
        self.step_m_var.set(config_data.get('step_m', defaults['step_m']))
        self.overlap_lat_var.set(config_data.get('overlap_lat', defaults['overlap_lat']))
        self.overlap_focus_var.set(config_data.get('overlap_focus', defaults['overlap_focus']))

        self.lens_profiles = config_data.get('lens_profiles', defaults['lens_profiles'])
        self.active_lens_var.set(config_data.get('active_lens', defaults['active_lens']))
        self.camera_profiles = config_data.get('camera_profiles', defaults['camera_profiles'])
        self.active_camera_var.set(config_data.get('active_camera', defaults['active_camera']))
        self.machine_profiles = config_data.get('machine_profiles', defaults['machine_profiles'])
        self.active_machine_var.set(config_data.get('active_machine', defaults['active_machine']))

        self.log_enabled_var.set(config_data.get('log_enabled', defaults['log_enabled']))
        self.log_dir_var.set(config_data.get('log_dir', defaults['log_dir']))

        self.layout = config_data.get('layout', defaults['layout'])
        self.saved_sashes = config_data.get('sashes', defaults['sashes'])

        for col in ["left", "center", "right"]:
            if col not in self.layout:
                self.layout[col] = []

        all_loaded = self.layout["left"] + self.layout["center"] + self.layout["right"]
        for required in ["cnc", "trigger", "log_panel", "hw", "mag", "seq", "grbl", "hist", "vis"]:
            if required not in all_loaded:
                # Force panel to appear if it was missing from the config
                if required == "log_panel":
                    self.layout["left"].insert(2, required)
                else:
                    self.layout["center"].insert(0, required)

    def _auto_connect_worker(self):
        time.sleep(1)
        self.root.after(0, self.update_com_ports, True)
        time.sleep(0.2)
        if self.com_port_var.get():
            self.root.after(0, self.toggle_connect_cnc, False)


if __name__ == '__main__':
    root = tk.Tk()
    try:
        root.iconbitmap(resource_path("logo.ico"))
    except Exception:
        pass
    app = CNCApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()