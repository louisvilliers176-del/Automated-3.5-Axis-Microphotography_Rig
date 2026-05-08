import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import shutil
import csv
import json
import time
from datetime import datetime
from pathlib import Path
import threading
import sys
import ctypes

try:
    import cv2
    import numpy as np
    HAS_CV2 = True
except ImportError:
    HAS_CV2 = False

try:
    import rawpy
    HAS_RAWPY = True
except ImportError:
    HAS_RAWPY = False

try:
    from PIL import Image as PILImage
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

try:
    import tifffile
    HAS_TIFFFILE = True
except ImportError:
    HAS_TIFFFILE = False

SETTINGS_FILE = Path(__file__).parent / "post_watcher_settings.json"

RAW_EXTS = {'.rw2', '.cr3', '.nef', '.dng', '.arw'}


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


class PostWatcherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Post-Watcher")
        self.root.geometry("720x780")
        self.root.minsize(660, 660)

        self.sd_dir_var = tk.StringVar()
        self.csv_path_var = tk.StringVar()
        self.out_dir_var = tk.StringVar()
        self.copy_raw_var = tk.BooleanVar(value=True)
        self.delete_source_var = tk.BooleanVar(value=False)

        # Preprocessing master toggle
        self.preprocess_var = tk.BooleanVar(value=False)

        # Independent per-step toggles
        self.apply_flat_field_var = tk.BooleanVar(value=True)
        self.apply_orb_var = tk.BooleanVar(value=True)
        self.apply_denoising_var = tk.BooleanVar(value=True)

        self.dust_map_path_var = tk.StringVar()
        self.dust_map_jpg_path_var = tk.StringVar()
        self.noise_profile_path_var = tk.StringVar()
        self.noise_profile_jpg_path_var = tk.StringVar()
        self.output_format_var = tk.StringVar(value="float32")

        self.progress_var = tk.DoubleVar(value=0.0)
        self.status_var = tk.StringVar(value="Waiting...")
        self.is_running = False

        self._load_settings()
        self._create_widgets()

    # --- PERSISTENCE ---

    def _load_settings(self):
        try:
            if SETTINGS_FILE.exists():
                with open(SETTINGS_FILE, 'r') as f:
                    s = json.load(f)
                self.sd_dir_var.set(s.get('sd_dir', ''))
                self.csv_path_var.set(s.get('csv_path', ''))
                self.out_dir_var.set(s.get('out_dir', ''))
                self.copy_raw_var.set(s.get('copy_raw', True))
                self.delete_source_var.set(s.get('delete_source', False))
                self.preprocess_var.set(s.get('preprocessing_enabled', False))
                self.apply_flat_field_var.set(s.get('apply_flat_field', True))
                self.apply_orb_var.set(s.get('apply_orb', True))
                self.apply_denoising_var.set(s.get('apply_denoising', True))
                self.dust_map_path_var.set(s.get('dust_map_path', ''))
                self.dust_map_jpg_path_var.set(s.get('dust_map_jpg_path', ''))
                self.noise_profile_path_var.set(s.get('noise_profile_path', ''))
                self.noise_profile_jpg_path_var.set(s.get('noise_profile_jpg_path', ''))
                self.output_format_var.set(s.get('output_format', 'float32'))
        except Exception:
            pass

    def _save_settings(self):
        try:
            s = {
                'sd_dir': self.sd_dir_var.get(),
                'csv_path': self.csv_path_var.get(),
                'out_dir': self.out_dir_var.get(),
                'copy_raw': self.copy_raw_var.get(),
                'delete_source': self.delete_source_var.get(),
                'preprocessing_enabled': self.preprocess_var.get(),
                'apply_flat_field': self.apply_flat_field_var.get(),
                'apply_orb': self.apply_orb_var.get(),
                'apply_denoising': self.apply_denoising_var.get(),
                'dust_map_path': self.dust_map_path_var.get(),
                'dust_map_jpg_path': self.dust_map_jpg_path_var.get(),
                'noise_profile_path': self.noise_profile_path_var.get(),
                'noise_profile_jpg_path': self.noise_profile_jpg_path_var.get(),
                'output_format': self.output_format_var.get(),
            }
            with open(SETTINGS_FILE, 'w') as f:
                json.dump(s, f, indent=2)
        except Exception:
            pass

    # --- UI ---

    def _create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 1. Ingestion Parameters
        path_lf = ttk.LabelFrame(main_frame, text="1. Ingestion Parameters", padding="10")
        path_lf.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(path_lf, text="SD Card Root Folder:").grid(row=0, column=0, sticky="w", pady=2)
        ttk.Entry(path_lf, textvariable=self.sd_dir_var).grid(row=0, column=1, sticky="ew", padx=5, pady=2)
        ttk.Button(path_lf, text="Browse", command=lambda: self._browse_dir(self.sd_dir_var)).grid(row=0, column=2, pady=2)

        ttk.Label(path_lf, text="CNC Log (.csv):").grid(row=1, column=0, sticky="w", pady=2)
        ttk.Entry(path_lf, textvariable=self.csv_path_var).grid(row=1, column=1, sticky="ew", padx=5, pady=2)
        ttk.Button(path_lf, text="Browse", command=self._browse_csv).grid(row=1, column=2, pady=2)

        ttk.Label(path_lf, text="Export Destination:").grid(row=2, column=0, sticky="w", pady=2)
        ttk.Entry(path_lf, textvariable=self.out_dir_var).grid(row=2, column=1, sticky="ew", padx=5, pady=2)
        ttk.Button(path_lf, text="Browse", command=lambda: self._browse_dir(self.out_dir_var)).grid(row=2, column=2, pady=2)
        path_lf.columnconfigure(1, weight=1)

        # 2. Preprocessing Pipeline
        prep_lf = ttk.LabelFrame(main_frame, text="2. Preprocessing Pipeline", padding="10")
        prep_lf.pack(fill=tk.X, pady=(0, 8))

        top_row = ttk.Frame(prep_lf)
        top_row.pack(fill=tk.X, pady=(0, 6))
        ttk.Checkbutton(top_row, text="Copy original files (archive)",
                        variable=self.copy_raw_var).pack(side=tk.LEFT, padx=(0, 20))
        ttk.Checkbutton(top_row, text="🗑 Delete sources after export",
                        variable=self.delete_source_var).pack(side=tk.LEFT, padx=(0, 20))

        can_preprocess = HAS_PIL or HAS_CV2
        self.preprocess_cb = ttk.Checkbutton(
            top_row,
            text="Enable preprocessing",
            variable=self.preprocess_var,
            command=self._on_preprocess_toggle,
            state=tk.NORMAL if can_preprocess else tk.DISABLED
        )
        self.preprocess_cb.pack(side=tk.LEFT)

        # Detail frame — shown when preprocessing is on
        self.prep_detail = ttk.Frame(prep_lf)
        self.prep_detail.pack(fill=tk.X, padx=(16, 0))

        # --- Flat field row ---
        ff_row = ttk.Frame(self.prep_detail)
        ff_row.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(2, 0))
        self.ff_cb = ttk.Checkbutton(
            ff_row, text="Correct Dust Map",
            variable=self.apply_flat_field_var,
            command=self._on_flat_field_toggle
        )
        self.ff_cb.pack(side=tk.LEFT)

        self.ff_entry_frame = ttk.Frame(self.prep_detail)
        self.ff_entry_frame.grid(row=1, column=0, columnspan=3, sticky="ew", padx=(20, 0), pady=(0, 4))
        ttk.Label(self.ff_entry_frame, text="Dust Map RAW (.tif):").grid(row=0, column=0, sticky="w")
        ttk.Entry(self.ff_entry_frame, textvariable=self.dust_map_path_var).grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(self.ff_entry_frame, text="Browse", command=self._browse_dustmap).grid(row=0, column=2)
        ttk.Label(self.ff_entry_frame, text="Dust Map JPG (.tif):").grid(row=1, column=0, sticky="w", pady=(2, 0))
        ttk.Entry(self.ff_entry_frame, textvariable=self.dust_map_jpg_path_var).grid(row=1, column=1, sticky="ew", padx=5, pady=(2, 0))
        ttk.Button(self.ff_entry_frame, text="Browse", command=self._browse_dustmap_jpg).grid(row=1, column=2, pady=(2, 0))
        self.ff_entry_frame.columnconfigure(1, weight=1)

        # --- ORB row ---
        orb_row = ttk.Frame(self.prep_detail)
        orb_row.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(2, 4))
        can_orb = HAS_CV2
        self.orb_cb = ttk.Checkbutton(
            orb_row, text="ORB Alignment (slower — requires opencv)",
            variable=self.apply_orb_var,
            state=tk.NORMAL if can_orb else tk.DISABLED
        )
        self.orb_cb.pack(side=tk.LEFT)
        if not can_orb:
            ttk.Label(orb_row, text="  [opencv missing]", foreground="orange").pack(side=tk.LEFT)

        # --- Denoising row ---
        dn_row = ttk.Frame(self.prep_detail)
        dn_row.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(2, 0))
        self.dn_cb = ttk.Checkbutton(
            dn_row, text="Anscombe Denoising",
            variable=self.apply_denoising_var,
            command=self._on_denoising_toggle
        )
        self.dn_cb.pack(side=tk.LEFT)

        self.dn_entry_frame = ttk.Frame(self.prep_detail)
        self.dn_entry_frame.grid(row=4, column=0, columnspan=3, sticky="ew", padx=(20, 0), pady=(0, 4))
        ttk.Label(self.dn_entry_frame, text="Noise Profile RAW (.json):").grid(row=0, column=0, sticky="w")
        ttk.Entry(self.dn_entry_frame, textvariable=self.noise_profile_path_var).grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(self.dn_entry_frame, text="Browse", command=self._browse_noise_profile).grid(row=0, column=2)
        ttk.Label(self.dn_entry_frame, text="Noise Profile JPG (.json):").grid(row=1, column=0, sticky="w", pady=(2, 0))
        ttk.Entry(self.dn_entry_frame, textvariable=self.noise_profile_jpg_path_var).grid(row=1, column=1, sticky="ew", padx=5, pady=(2, 0))
        ttk.Button(self.dn_entry_frame, text="Browse", command=self._browse_noise_profile_jpg).grid(row=1, column=2, pady=(2, 0))
        self.dn_entry_frame.columnconfigure(1, weight=1)

        # --- Output format ---
        fmt_frame = ttk.Frame(self.prep_detail)
        fmt_frame.grid(row=5, column=0, columnspan=3, sticky="w", pady=(6, 2))
        ttk.Label(fmt_frame, text="Output format:").pack(side=tk.LEFT)
        ttk.Radiobutton(fmt_frame, text="float32 linear (ML pipeline)",
                        variable=self.output_format_var, value="float32").pack(side=tk.LEFT, padx=(8, 0))
        ttk.Radiobutton(fmt_frame, text="uint16 gamma sRGB (Helicon / Zerene)",
                        variable=self.output_format_var, value="uint16_gamma").pack(side=tk.LEFT, padx=(8, 0))

        self.prep_detail.columnconfigure(1, weight=1)

        if not can_preprocess:
            ttk.Label(self.prep_detail,
                      text="⚠  Missing modules: PIL and opencv — preprocessing unavailable",
                      foreground="orange").grid(row=6, column=0, columnspan=3, pady=4)

        self._on_preprocess_toggle()

        # 3. Action
        ttk.Button(main_frame, text="🚀  Start Extraction & Sorting",
                   command=self.start_processing).pack(fill=tk.X, ipady=5, pady=(0, 8))

        # 4. Progress & Logs
        prog_lf = ttk.LabelFrame(main_frame, text="3. Progress & Logs", padding="10")
        prog_lf.pack(fill=tk.BOTH, expand=True)

        self.progress_bar = ttk.Progressbar(prog_lf, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=5)
        ttk.Label(prog_lf, textvariable=self.status_var, font=("Arial", 9, "bold")).pack(anchor="w", pady=2)

        self.log_text = tk.Text(prog_lf, wrap=tk.WORD, state=tk.DISABLED, height=12)
        scrollbar = ttk.Scrollbar(prog_lf, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.config(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.log_text.tag_configure("error", foreground="red")
        self.log_text.tag_configure("success", foreground="green")
        self.log_text.tag_configure("info", foreground="blue")
        self.log_text.tag_configure("warning", foreground="orange")

    def _on_preprocess_toggle(self):
        enabled = self.preprocess_var.get()
        state = tk.NORMAL if enabled else tk.DISABLED
        for cb in (self.ff_cb, self.orb_cb, self.dn_cb):
            try:
                cb.config(state=state)
            except tk.TclError:
                pass
        if enabled:
            self._on_flat_field_toggle()
            self._on_denoising_toggle()
        else:
            self._set_frame_state(self.ff_entry_frame, tk.DISABLED)
            self._set_frame_state(self.dn_entry_frame, tk.DISABLED)

    def _on_flat_field_toggle(self):
        state = tk.NORMAL if (self.preprocess_var.get() and self.apply_flat_field_var.get()) else tk.DISABLED
        self._set_frame_state(self.ff_entry_frame, state)

    def _on_denoising_toggle(self):
        state = tk.NORMAL if (self.preprocess_var.get() and self.apply_denoising_var.get()) else tk.DISABLED
        self._set_frame_state(self.dn_entry_frame, state)

    @staticmethod
    def _set_frame_state(frame, state):
        for child in frame.winfo_children():
            try:
                child.config(state=state)
            except tk.TclError:
                pass

    # --- BROWSE CALLBACKS ---

    def _browse_dir(self, var):
        d = filedialog.askdirectory()
        if d: var.set(d)

    def _browse_csv(self):
        f = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if f: self.csv_path_var.set(f)

    def _browse_dustmap(self):
        f = filedialog.askopenfilename(filetypes=[("TIFF Files", "*.tif *.tiff")])
        if f: self.dust_map_path_var.set(f)

    def _browse_dustmap_jpg(self):
        f = filedialog.askopenfilename(filetypes=[("TIFF Files", "*.tif *.tiff")])
        if f: self.dust_map_jpg_path_var.set(f)

    def _browse_noise_profile(self):
        f = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")])
        if f: self.noise_profile_path_var.set(f)

    def _browse_noise_profile_jpg(self):
        f = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")])
        if f: self.noise_profile_jpg_path_var.set(f)

    # --- LOGGING ---

    def log(self, message, level="normal"):
        timestamp = time.strftime("%H:%M:%S")
        self.root.after(0, self._append_log, f"[{timestamp}] {message}\n", level)

    def _append_log(self, text, level):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, text, level)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    # --- START ---

    def start_processing(self):
        if self.is_running:
            return

        sd_path = self.sd_dir_var.get()
        csv_path = self.csv_path_var.get()
        out_path = self.out_dir_var.get()

        if not all([sd_path, csv_path, out_path]):
            messagebox.showerror("Error", "Please fill in all paths.")
            return

        if self.preprocess_var.get():
            if self.apply_flat_field_var.get():
                if not self.dust_map_path_var.get() and not self.dust_map_jpg_path_var.get():
                    messagebox.showerror("Error", "Dust Map correction enabled but no file selected (RAW or JPG).")
                    return
            if self.apply_denoising_var.get():
                if not self.noise_profile_path_var.get() and not self.noise_profile_jpg_path_var.get():
                    messagebox.showerror("Error", "Denoising enabled but no noise profile selected (RAW or JPG).")
                    return

        delete_source = self.delete_source_var.get()
        if delete_source:
            if not messagebox.askyesno(
                "Delete confirmation",
                "Source files will be permanently deleted after export.\n\nContinue?",
                icon="warning"
            ):
                return

        self._save_settings()
        self.is_running = True
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete('1.0', tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.progress_var.set(0)

        threading.Thread(
            target=self._worker_thread,
            args=(sd_path, csv_path, out_path,
                  self.copy_raw_var.get(),
                  delete_source,
                  self.preprocess_var.get(),
                  self.apply_flat_field_var.get(),
                  self.apply_orb_var.get(),
                  self.apply_denoising_var.get(),
                  self.dust_map_path_var.get(),
                  self.dust_map_jpg_path_var.get(),
                  self.noise_profile_path_var.get(),
                  self.noise_profile_jpg_path_var.get(),
                  self.output_format_var.get()),
            daemon=True
        ).start()

    # --- PREPROCESSING HELPERS (static) ---

    @staticmethod
    def _load_raw_linear(path):
        with rawpy.imread(str(path)) as raw:
            rgb16 = raw.postprocess(
                gamma=(1, 1),
                no_auto_bright=True,
                use_camera_wb=True,
                output_bps=16
            )
        return rgb16.astype(np.float32) / 65535.0

    @staticmethod
    def _load_any_image(path):
        """Load RAW or standard image (JPG/PNG/TIFF) as float32 [0,1] RGB."""
        suffix = Path(path).suffix.lower()
        if suffix in RAW_EXTS and HAS_RAWPY:
            return PostWatcherApp._load_raw_linear(path)
        if HAS_PIL:
            img = PILImage.open(str(path)).convert('RGB')
            return np.array(img).astype(np.float32) / 255.0
        if HAS_CV2:
            img = cv2.imread(str(path), cv2.IMREAD_COLOR)
            img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
            return img.astype(np.float32) / 255.0
        raise RuntimeError("Neither PIL nor OpenCV available to load the image.")

    @staticmethod
    def _load_dust_map(path):
        if HAS_TIFFFILE:
            dm = tifffile.imread(str(path)).astype(np.float32)
        else:
            dm = np.array(PILImage.open(str(path))).astype(np.float32)
        if dm.max() > 1.1:
            dm = dm / 65535.0
        return dm

    @staticmethod
    def _apply_flat_field(img, dust_map, img_shape_hw):
        dm = dust_map
        if dm.shape[:2] != img_shape_hw:
            dm = cv2.resize(dm, (img_shape_hw[1], img_shape_hw[0]))
        mean = dm.mean()
        if mean <= 0:
            return img
        flat_norm = dm / mean
        if flat_norm.ndim == 2:
            flat_norm = flat_norm[:, :, np.newaxis]
        return np.clip(img / flat_norm, 0.0, 1.0).astype(np.float32)

    # Laplacian variance threshold below which a frame is considered too blurry
    # for ORB features to be reliable. Empirical value — can be adjusted.
    ORB_BLUR_THRESHOLD = 50.0

    @staticmethod
    def _compute_orb_transforms(src_files):
        """Pass 1: computes consecutive alignment matrices (works on RAW and JPG/PNG)."""
        N = len(src_files)
        if N <= 1:
            return [np.eye(3)] * max(0, N - 1)

        orb = cv2.ORB_create(nfeatures=2000)
        matcher = cv2.BFMatcher(cv2.NORM_HAMMING, crossCheck=True)
        transforms = [np.eye(3)] * (N - 1)

        prev_gray = None
        prev_sharp = False
        for i, path in enumerate(src_files):
            img = PostWatcherApp._load_any_image(path)
            u8 = (img * 255).clip(0, 255).astype(np.uint8)
            gray = cv2.cvtColor(u8, cv2.COLOR_RGB2GRAY)
            del img, u8

            lap_var = cv2.Laplacian(gray, cv2.CV_64F).var()
            is_sharp = lap_var >= PostWatcherApp.ORB_BLUR_THRESHOLD

            if prev_gray is not None and is_sharp and prev_sharp:
                j = i - 1
                kp1, des1 = orb.detectAndCompute(prev_gray, None)
                kp2, des2 = orb.detectAndCompute(gray, None)
                if (des1 is not None and des2 is not None
                        and len(des1) >= 4 and len(des2) >= 4):
                    matches = sorted(matcher.match(des1, des2), key=lambda x: x.distance)
                    good = matches[:max(4, int(len(matches) * 0.2))]
                    if len(good) >= 4:
                        pts1 = np.float32([kp1[m.queryIdx].pt for m in good]).reshape(-1, 1, 2)
                        pts2 = np.float32([kp2[m.trainIdx].pt for m in good]).reshape(-1, 1, 2)
                        M, _ = cv2.estimateAffinePartial2D(pts1, pts2, method=cv2.RANSAC)
                        if M is not None:
                            transforms[j] = np.vstack([M, [0, 0, 1]])

            prev_gray = gray
            prev_sharp = is_sharp

        return transforms

    @staticmethod
    def _get_cumulative_transform(transforms, i, ref_idx):
        M = np.eye(3)
        if i < ref_idx:
            for j in range(i, ref_idx):
                M = transforms[j] @ M
        elif i > ref_idx:
            for j in range(i - 1, ref_idx - 1, -1):
                M = np.linalg.inv(transforms[j]) @ M
        return M

    @staticmethod
    def _anscombe_denoise(img, photon_scale):
        t = 2.0 * np.sqrt(img * photon_scale + 3.0 / 8.0)
        t = cv2.GaussianBlur(t, (0, 0), sigmaX=1.0)
        return np.clip(((t / 2.0) ** 2 - 3.0 / 8.0) / photon_scale, 0.0, 1.0).astype(np.float32)

    @staticmethod
    def _save_tiff(path, img_float32, output_format="float32"):
        if output_format == "uint16_gamma":
            gamma = np.power(np.clip(img_float32, 0.0, 1.0), 1.0 / 2.2)
            u16 = (gamma * 65535).clip(0, 65535).astype(np.uint16)
            if HAS_TIFFFILE:
                tifffile.imwrite(str(path), u16, photometric='rgb')
            elif HAS_PIL:
                PILImage.fromarray(u16).save(str(path))
            else:
                raise RuntimeError("No TIFF module available (install tifffile).")
        else:
            if HAS_TIFFFILE:
                tifffile.imwrite(str(path), img_float32, photometric='rgb')
            elif HAS_PIL:
                u8 = (img_float32 * 255).clip(0, 255).astype(np.uint8)
                PILImage.fromarray(u8).save(str(path))
            else:
                raise RuntimeError("No TIFF module available (install tifffile).")

    # --- WORKER ---

    def _worker_thread(self, sd_dir, log_file, output_dir, copy_raw, delete_source,
                       preprocess, apply_flat_field, apply_orb, apply_denoising,
                       dust_map_path, dust_map_jpg_path,
                       noise_profile_path, noise_profile_jpg_path,
                       output_format="float32"):
        try:
            dust_map = None
            photon_scale = None
            had_processing_errors = False

            if preprocess and not HAS_TIFFFILE:
                self.log("⚠  tifffile not installed — TIFF float32 export unavailable, fallback uint8.", "warning")

            # Parse CSV
            self.log("Parsing CNC CSV file...", "info")
            cnc_data = []
            metadata_header = {}

            with open(log_file, mode='r', encoding='utf-8') as f:
                lines = f.readlines()

            data_lines = []
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                if line.startswith('#'):
                    parts = line.split(':', 1)
                    if len(parts) == 2:
                        metadata_header[parts[0].replace('#', '').strip()] = parts[1].strip()
                else:
                    data_lines.append(line)

            if len(data_lines) < 2:
                raise ValueError("The CSV file contains no position data.")

            for row in csv.DictReader(data_lines):
                cnc_data.append({
                    'frame_index': int(row['frame_index']),
                    'timestamp': float(row['timestamp']),
                    'x_um': float(row['x_um']),
                    'y_um': float(row['y_um']),
                    'z_um': float(row['z_um'])
                })

            n_expected = len(cnc_data)
            self.log(f"{n_expected} positions found in CNC log.")

            y_values = set(d['y_um'] for d in cnc_data)
            z_values = set(d['z_um'] for d in cnc_data)
            self.log(f"Panorama: grid {len(y_values)} (Y) × {len(z_values)} (Z).")

            # Scan source folder
            self.root.after(0, self.status_var.set, "Scanning source folder...")
            self.log("Scanning source folder...", "info")
            sd_path = Path(sd_dir)
            valid_exts = {'.rw2', '.jpg', '.jpeg', '.cr3', '.tif', '.tiff', '.nef', '.dng', '.arw', '.png'}
            src_files = sorted(
                [f for f in sd_path.rglob('*') if f.is_file() and f.suffix.lower() in valid_exts],
                key=lambda x: x.stat().st_mtime
            )

            if len(src_files) < n_expected:
                raise ValueError(f"{len(src_files)} photos found, {n_expected} expected.")

            stack_files = src_files[-n_expected:]
            older = len(src_files) - n_expected
            if older > 0:
                self.log(f"{older} older photos ignored.", "warning")

            # Log file type summary
            exts_found = {}
            for f in stack_files:
                ext = f.suffix.lower()
                exts_found[ext] = exts_found.get(ext, 0) + 1
            self.log("Detected types: " + ", ".join(f"{v}× {k}" for k, v in sorted(exts_found.items())))

            # Determine dominant image type (RAW vs JPG) and select appropriate profiles
            jpg_exts = {'.jpg', '.jpeg'}
            n_raw = sum(v for k, v in exts_found.items() if k in RAW_EXTS)
            n_jpg = sum(v for k, v in exts_found.items() if k in jpg_exts)
            stack_is_jpg = n_jpg > n_raw
            if n_raw > 0 and n_jpg > 0:
                self.log(f"⚠  Mixed stack ({n_raw} RAW, {n_jpg} JPG) — {'JPG' if stack_is_jpg else 'RAW'} profile selected (dominant).", "warning")
            elif stack_is_jpg:
                self.log("Stack type: JPG — JPG profiles used.", "info")
            else:
                self.log("Stack type: RAW — RAW profiles used.", "info")

            if preprocess:
                if apply_flat_field:
                    active_dm_path = (dust_map_jpg_path if stack_is_jpg else dust_map_path) or dust_map_path or dust_map_jpg_path
                    if not active_dm_path:
                        raise ValueError("Dust Map correction enabled but no file selected.")
                    self.log(f"Loading {'JPG' if stack_is_jpg else 'RAW'} Dust Map...", "info")
                    dust_map = self._load_dust_map(active_dm_path)
                    self.log(f"Dust map loaded ({dust_map.shape[1]}×{dust_map.shape[0]}).", "success")

                if apply_denoising:
                    active_np_path = (noise_profile_jpg_path if stack_is_jpg else noise_profile_path) or noise_profile_path or noise_profile_jpg_path
                    if not active_np_path:
                        raise ValueError("Denoising enabled but no noise profile selected.")
                    self.log(f"Loading {'JPG' if stack_is_jpg else 'RAW'} noise profile...", "info")
                    with open(active_np_path, 'r') as f:
                        noise_data = json.load(f)
                    photon_scale = noise_data.get('poisson_calibration', noise_data).get('photon_scale')
                    if photon_scale is None:
                        raise ValueError("photon_scale not found in noise profile.")
                    self.log(f"photon_scale = {photon_scale:.1f}", "success")

            time_sd = stack_files[-1].stat().st_mtime - stack_files[0].stat().st_mtime
            time_csv = cnc_data[-1]['timestamp'] - cnc_data[0]['timestamp']
            if abs(time_sd - time_csv) > 25.0:
                self.log(f"⚠  Time gap (Log: {time_csv:.1f}s | SD: {time_sd:.1f}s).", "error")

            # Build export tree
            now = datetime.now()
            root_folder_name = f"{now.strftime('%y%m%d_%H%M')}_{n_expected}pics_{len(y_values)}x{len(z_values)}"
            export_root = Path(output_dir) / root_folder_name
            export_root.mkdir(parents=True, exist_ok=True)
            self.log(f"Folder created: {root_folder_name}", "success")

            # Phase 1: sort and copy + JSON sidecars
            stacks_map = {}

            phase1_total = n_expected
            for i, (entry, src_file) in enumerate(zip(cnc_data, stack_files)):
                y_val, z_val = entry['y_um'], entry['z_um']
                folder_name = f"Y{y_val:.1f}_Z{z_val:.1f}"
                stack_dir = export_root / folder_name
                stack_dir.mkdir(exist_ok=True)

                new_name = f"img_{entry['frame_index']:04d}_X{entry['x_um']:.1f}{src_file.suffix.lower()}"
                dest_img = stack_dir / new_name
                dest_json = stack_dir / f"img_{entry['frame_index']:04d}_X{entry['x_um']:.1f}.json"

                self.root.after(0, self.status_var.set, f"Copy: {new_name}")

                if copy_raw:
                    shutil.copy2(src_file, dest_img)

                with open(dest_json, 'w', encoding='utf-8') as jf:
                    json.dump({
                        'frame_index': entry['frame_index'],
                        'x_um': entry['x_um'],
                        'y_um': entry['y_um'],
                        'z_um': entry['z_um'],
                        'original_timestamp': entry['timestamp'],
                        'metadata_context': metadata_header
                    }, jf, indent=2)

                if folder_name not in stacks_map:
                    stacks_map[folder_name] = {'dir': stack_dir, 'src_files': []}
                stacks_map[folder_name]['src_files'].append(src_file)

                prog = ((i + 1) / phase1_total) * (50.0 if preprocess else 100.0)
                self.root.after(0, self.progress_var.set, prog)

            if not copy_raw:
                self.log("Original files not copied (option disabled).", "warning")

            # Phase 2: preprocessing per stack
            if preprocess:
                steps = []
                if apply_flat_field: steps.append("flat field")
                if apply_orb and HAS_CV2: steps.append("ORB")
                if apply_denoising: steps.append("Anscombe")
                self.log(f"--- Phase 2: Preprocessing [{', '.join(steps)}] ---", "info")

                stack_list = sorted(stacks_map.items())
                n_stacks = len(stack_list)

                for k, (folder_name, info) in enumerate(stack_list):
                    stack_dir = info['dir']
                    src_list = sorted(info['src_files'], key=lambda p: p.name)
                    self.root.after(0, self.status_var.set, f"Processing {k+1}/{n_stacks}: {folder_name}")
                    self.log(f"Stack {k+1}/{n_stacks}: {folder_name} ({len(src_list)} frames)")

                    try:
                        N = len(src_list)
                        ref_idx = N // 2

                        # Pass 1 ORB (optional)
                        transforms = None
                        if apply_orb and HAS_CV2 and N > 1:
                            self.log(f"  ORB alignment — pass 1 ({N} frames)...", "info")
                            transforms = self._compute_orb_transforms(src_list)

                        # Dimensions from reference frame
                        ref_img = self._load_any_image(src_list[ref_idx])
                        H, W = ref_img.shape[:2]
                        del ref_img

                        # Pass 2: streaming frame by frame
                        self.log("  Processing frame by frame (streaming)...", "info")
                        processed_dir = stack_dir / "processed"
                        processed_dir.mkdir(exist_ok=True)

                        for fi, src in enumerate(src_list):
                            frame = self._load_any_image(src)

                            if apply_flat_field and dust_map is not None:
                                frame = self._apply_flat_field(frame, dust_map, (H, W))

                            if transforms is not None:
                                cum_M = self._get_cumulative_transform(transforms, fi, ref_idx)
                                if not np.allclose(cum_M, np.eye(3)):
                                    frame = cv2.warpAffine(frame, cum_M[:2, :], (W, H),
                                                           flags=cv2.INTER_LINEAR,
                                                           borderMode=cv2.BORDER_REPLICATE)

                            if apply_denoising and photon_scale is not None:
                                frame = self._anscombe_denoise(frame, photon_scale)

                            self._save_tiff(processed_dir / (src.stem + ".tif"), frame, output_format)
                            del frame

                            inner_prog = 50.0 + (k / n_stacks + (fi + 1) / (N * n_stacks)) * 50.0
                            self.root.after(0, self.progress_var.set, inner_prog)

                        self.log(f"  ✅ {N} frames → processed/", "success")

                    except OSError as e:
                        had_processing_errors = True
                        import errno
                        if e.errno == errno.ENOSPC or "requested and 0 written" in str(e):
                            self.log(f"  ✖ Destination disk full — cannot write ({folder_name}). Free up space and retry.", "error")
                        else:
                            self.log(f"  ✖ Disk error on {folder_name}: {e}", "error")
                    except Exception as e:
                        had_processing_errors = True
                        self.log(f"  ✖ Error on {folder_name}: {e}", "error")
                        prog = 50.0 + ((k + 1) / n_stacks) * 50.0
                        self.root.after(0, self.progress_var.set, prog)

            if had_processing_errors:
                self.log("⚠  Preprocessing completed with errors — see ✖ lines above.", "warning")
            else:
                self.log("✅ Extraction and sorting completed successfully!", "success")

            if delete_source and had_processing_errors:
                self.log("🛑 Deletion cancelled: preprocessing errors were detected. Verify exports before deleting sources.", "error")
            elif delete_source:
                self.log(f"Deleting {len(stack_files)} source files...", "info")
                deleted, errors = 0, 0
                for f in stack_files:
                    try:
                        f.unlink()
                        deleted += 1
                    except Exception as e:
                        self.log(f"  ✖ Cannot delete {f.name}: {e}", "error")
                        errors += 1
                if errors == 0:
                    self.log(f"🗑 {deleted} source files deleted.", "success")
                else:
                    self.log(f"🗑 {deleted} deleted, {errors} errors.", "warning")

            self.root.after(0, self.status_var.set, "Done.")

        except Exception as e:
            self.log(f"CRITICAL ERROR: {e}", "error")
            self.root.after(0, self.status_var.set, "Error.")
        finally:
            self.root.after(0, self._cleanup)

    def _cleanup(self):
        self.is_running = False


# --- LAUNCHER ---
if __name__ == "__main__":
    if os.name == 'nt':
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('microphoto.post_watcher.v4')
        except Exception:
            pass

    root = tk.Tk()
    try:
        root.iconbitmap(resource_path("logo.ico"))
    except Exception:
        pass

    app = PostWatcherApp(root)
    root.mainloop()
