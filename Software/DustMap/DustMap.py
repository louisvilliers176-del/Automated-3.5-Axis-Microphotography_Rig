import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import cv2
import numpy as np
import os
import json
from datetime import date
from PIL import Image, ImageTk

try:
    import rawpy
except ImportError:
    rawpy = None

try:
    import tifffile
except ImportError:
    tifffile = None


VALID_IMG_EXTS = {'.rw2', '.cr3', '.nef', '.dng', '.arw', '.jpg', '.jpeg'}


class DustMapProfiler:
    def __init__(self, root):
        self.root = root
        self.root.title("Phase 0.4: Dust Map Profiler — Flat Field Calibration")
        self.root.minsize(1000, 600)

        self.dust_map = None        # float32 H×W×3, values 0–65535
        self.dust_map_path = None
        self._flat_files = []
        self.production_raw = None  # float32 H×W×3, values 0–65535
        self.tk_img = None

        self.display_mode = tk.StringVar(value="dustmap")

        self._setup_ui()
        self.canvas.bind("<Configure>", lambda _: self._refresh_canvas())

    def _setup_ui(self):
        ctrl = ttk.Frame(self.root, padding=10)
        ctrl.pack(side=tk.LEFT, fill=tk.Y)

        ttk.Label(ctrl, text="Phase 0.4 — Dust Map Profiler",
                  font=("Arial", 11, "bold")).pack(pady=(0, 10))

        # --- 1. Flat frames ---
        build_frame = ttk.LabelFrame(ctrl, text="1. Flat Frames → Dust Map", padding=8)
        build_frame.pack(fill=tk.X, pady=6)

        ttk.Button(build_frame, text="Select flat frames folder (RAW / JPG)",
                   command=self._load_flat_frames).pack(fill=tk.X, pady=2)

        self.lbl_frames = ttk.Label(build_frame, text="Flat frames: 0")
        self.lbl_frames.pack(anchor="w", pady=(2, 6))

        ttk.Button(build_frame, text="Build Dust Map",
                   command=self._build_dust_map).pack(fill=tk.X, pady=2)

        self.lbl_mean = ttk.Label(build_frame, text="Mean brightness: --")
        self.lbl_mean.pack(anchor="w", pady=(6, 0))
        self.lbl_variation = ttk.Label(build_frame, text="Center variation: --")
        self.lbl_variation.pack(anchor="w")
        self.lbl_go = ttk.Label(build_frame, text="GO Criterion: --",
                                font=("Arial", 10, "bold"))
        self.lbl_go.pack(anchor="w", pady=(4, 0))

        ttk.Label(build_frame,
                  text="No dark subtraction — CMOS glow\nat short exposure is below noise threshold.",
                  font=("Arial", 8), foreground="gray").pack(anchor="w", pady=(6, 0))

        # --- Metadata ---
        meta_frame = ttk.LabelFrame(ctrl, text="Metadata", padding=8)
        meta_frame.pack(fill=tk.X, pady=6)

        row = ttk.Frame(meta_frame)
        row.pack(fill=tk.X, pady=2)
        ttk.Label(row, text="Objective:").pack(side=tk.LEFT)
        self.entry_objective = ttk.Entry(row)
        self.entry_objective.insert(0, "Nikon M Plan 10x")
        self.entry_objective.pack(side=tk.RIGHT, fill=tk.X, expand=True)

        row2 = ttk.Frame(meta_frame)
        row2.pack(fill=tk.X, pady=2)
        ttk.Label(row2, text="ISO :").pack(side=tk.LEFT)
        self.spin_iso = ttk.Spinbox(row2, from_=50, to=3200, increment=100, width=6)
        self.spin_iso.set(100)
        self.spin_iso.pack(side=tk.RIGHT)

        row3 = ttk.Frame(meta_frame)
        row3.pack(fill=tk.X, pady=2)
        ttk.Label(row3, text="Expo flats :").pack(side=tk.LEFT)
        self.entry_expo_flat = ttk.Entry(row3, width=8)
        self.entry_expo_flat.insert(0, "1/60")
        self.entry_expo_flat.pack(side=tk.RIGHT)

        row4 = ttk.Frame(meta_frame)
        row4.pack(fill=tk.X, pady=2)
        ttk.Label(row4, text="Expo production :").pack(side=tk.LEFT)
        self.entry_expo_prod = ttk.Entry(row4, width=8)
        self.entry_expo_prod.insert(0, "1/4")
        self.entry_expo_prod.pack(side=tk.RIGHT)

        ttk.Label(meta_frame,
                  text="Same ISO — different exposures OK\n(flat_norm is exposure-independent)",
                  font=("Arial", 8), foreground="gray").pack(anchor="w", pady=(4, 0))

        # --- 2. Save ---
        save_frame = ttk.LabelFrame(ctrl, text="2. Save", padding=8)
        save_frame.pack(fill=tk.X, pady=6)

        ttk.Button(save_frame, text="💾 Save .tif",
                   command=self._save_tif).pack(fill=tk.X, pady=2)
        ttk.Button(save_frame, text="⚙️ Save to config.json",
                   command=self._write_config).pack(fill=tk.X, pady=2)

        self.lbl_save_status = ttk.Label(save_frame, text="")
        self.lbl_save_status.pack(anchor="w", pady=(4, 0))

        ttk.Separator(ctrl, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)

        # --- 3. Validation ---
        val_frame = ttk.LabelFrame(ctrl, text="3. Validation", padding=8)
        val_frame.pack(fill=tk.X, pady=6)

        ttk.Button(val_frame, text="Load production frame (RAW / JPG)",
                   command=self._load_production_frame).pack(fill=tk.X, pady=2)

        self.lbl_prod = ttk.Label(val_frame, text="No frame loaded",
                                  font=("Arial", 8), foreground="gray")
        self.lbl_prod.pack(anchor="w", pady=2)

        # --- Display ---
        ttk.Separator(ctrl, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        ttk.Label(ctrl, text="Display:").pack(anchor="w")

        for text, value in [
            ("Dust Map (false colors)", "dustmap"),
            ("Original frame", "before"),
            ("Corrected frame", "after"),
        ]:
            ttk.Radiobutton(ctrl, text=text, variable=self.display_mode,
                            value=value, command=self._refresh_canvas).pack(anchor="w")

        # --- Canvas ---
        self.canvas = tk.Canvas(self.root, bg="gray20")
        self.canvas.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

    # ------------------------------------------------------------------
    # RAW Loading
    # ------------------------------------------------------------------

    def _load_raw_linear(self, path):
        ext = os.path.splitext(path)[1].lower()
        if ext in ('.jpg', '.jpeg'):
            img = np.array(Image.open(path).convert('RGB'), dtype=np.float32)
            return img * 257.0  # 0-255 → 0-65535
        if rawpy is None:
            messagebox.showerror("Error", "Install rawpy: pip install rawpy")
            return None
        try:
            with rawpy.imread(path) as raw:
                rgb16 = raw.postprocess(
                    gamma=(1, 1),
                    no_auto_bright=True,
                    use_camera_wb=True,
                    output_bps=16,
                )
            return rgb16.astype(np.float32)
        except Exception as e:
            messagebox.showerror("RAW Read Error",
                                 f"Cannot read:\n{os.path.basename(path)}\n\n{e}")
            return None

    def _list_raw_files(self, dir_path):
        return sorted([
            os.path.join(dir_path, f)
            for f in os.listdir(dir_path)
            if os.path.splitext(f)[1].lower() in VALID_IMG_EXTS
        ])

    # ------------------------------------------------------------------
    # Flat frames → Dust Map
    # ------------------------------------------------------------------

    def _load_flat_frames(self):
        dir_path = filedialog.askdirectory(title="Flat frames folder RAW / JPG")
        if not dir_path:
            return
        self._flat_files = self._list_raw_files(dir_path)
        self.lbl_frames.config(text=f"Flat frames: {len(self._flat_files)}")
        if len(self._flat_files) < 3:
            messagebox.showwarning("Warning",
                                   f"{len(self._flat_files)} file(s). "
                                   "5–10 frames recommended.")

    def _build_dust_map(self):
        if len(self._flat_files) < 3:
            messagebox.showerror("Error", "Load the flat frames folder first (min. 3).")
            return

        frames = []
        for i, f in enumerate(self._flat_files):
            self.lbl_mean.config(text=f"Loading {i + 1}/{len(self._flat_files)}…")
            self.root.update()
            img = self._load_raw_linear(f)
            if img is not None:
                frames.append(img)

        if len(frames) < 3:
            messagebox.showerror("Error", "Not enough readable RAW frames.")
            return

        stack = np.stack(frames, axis=0)
        self.dust_map = np.median(stack, axis=0)  # float32, 0–65535

        # Stats — zone centrale 500×500, canal vert
        H, W = self.dust_map.shape[:2]
        cy, cx = H // 2, W // 2
        half = 250
        center_g = self.dust_map[cy - half:cy + half, cx - half:cx + half, 1]

        mean_val = self.dust_map[:, :, 1].mean()
        center_var_pct = (center_g.std() / (center_g.mean() + 1e-6)) * 100

        self.lbl_mean.config(text=f"Mean brightness: {mean_val / 65535:.4f}")
        self.lbl_variation.config(text=f"Center variation: {center_var_pct:.3f} %")

        if center_var_pct < 0.5:
            self.lbl_go.config(text="GO Criterion: ✓ GO  (< 0.5 %)", foreground="green")
        else:
            self.lbl_go.config(
                text=f"GO Criterion: ✗ NO GO ({center_var_pct:.3f} % > 0.5 %)",
                foreground="red",
            )

        self.display_mode.set("dustmap")
        self._refresh_canvas()

    # ------------------------------------------------------------------
    # Validation
    # ------------------------------------------------------------------

    def _load_production_frame(self):
        path = filedialog.askopenfilename(
            title="Production frame (RAW / JPG)",
            filetypes=[("Images", "*.rw2 *.cr3 *.nef *.dng *.arw *.jpg *.jpeg"),
                       ("RAW", "*.rw2 *.cr3 *.nef *.dng *.arw"),
                       ("JPEG", "*.jpg *.jpeg")],
        )
        if not path:
            return
        img = self._load_raw_linear(path)
        if img is None:
            return
        self.production_raw = img
        self.lbl_prod.config(text=os.path.basename(path), foreground="")
        self.display_mode.set("before")
        self._refresh_canvas()

    # ------------------------------------------------------------------
    # Affichage
    # ------------------------------------------------------------------

    def _to_display_rgb(self, img_float32):
        """float32 [0–65535] → uint8 RGB gamma-corrected for display."""
        normalized = (img_float32 / 65535.0).clip(0, 1)
        gamma = np.power(normalized, 1 / 2.2)
        return (gamma * 255).astype(np.uint8)

    def _get_frame_to_show(self):
        mode = self.display_mode.get()

        if mode == "dustmap":
            if self.dust_map is None:
                return None
            green = self.dust_map[:, :, 1]
            p1, p99 = np.percentile(green, [1, 99])
            stretched = ((green - p1) / (p99 - p1 + 1e-6) * 255).clip(0, 255).astype(np.uint8)
            colored = cv2.applyColorMap(stretched, cv2.COLORMAP_INFERNO)
            return cv2.cvtColor(colored, cv2.COLOR_BGR2RGB)

        if mode == "before":
            if self.production_raw is None:
                return None
            return self._to_display_rgb(self.production_raw)

        if mode == "after":
            if self.production_raw is None or self.dust_map is None:
                return None
            if self.production_raw.shape != self.dust_map.shape:
                messagebox.showerror("Error",
                                     "Incompatible resolutions between frame and Dust Map.")
                return None
            flat_norm = np.ones_like(self.dust_map)
            for c in range(3):
                ch_mean = self.dust_map[:, :, c].mean()
                flat_norm[:, :, c] = self.dust_map[:, :, c] / (ch_mean + 1e-6)
            corrected = (self.production_raw / (flat_norm + 1e-6)).clip(0, 65535)
            return self._to_display_rgb(corrected)

        return None

    def _refresh_canvas(self):
        arr = self._get_frame_to_show()
        if arr is None:
            return

        pil_img = Image.fromarray(arr)
        cw = self.canvas.winfo_width() or 900
        ch = self.canvas.winfo_height() or 700
        scale = min(cw / pil_img.width, ch / pil_img.height, 1.0)
        new_size = (max(1, int(pil_img.width * scale)), max(1, int(pil_img.height * scale)))

        self.tk_img = ImageTk.PhotoImage(pil_img.resize(new_size, Image.Resampling.LANCZOS))
        self.canvas.delete("all")
        self.canvas.create_image(cw // 2, ch // 2, anchor=tk.CENTER, image=self.tk_img)

    # ------------------------------------------------------------------
    # Save
    # ------------------------------------------------------------------

    def _save_tif(self):
        if self.dust_map is None:
            messagebox.showerror("Error", "Build the Dust Map first.")
            return

        obj_name = self.entry_objective.get().replace(" ", "_")
        default_name = f"dust_map_{obj_name}.tif"
        path = filedialog.asksaveasfilename(
            defaultextension=".tif",
            initialfile=default_name,
            filetypes=[("TIFF", "*.tif *.tiff")],
        )
        if not path:
            return

        if tifffile is not None:
            tifffile.imwrite(path, self.dust_map)
        else:
            img_u16 = self.dust_map.clip(0, 65535).astype(np.uint16)
            cv2.imwrite(path, cv2.cvtColor(img_u16, cv2.COLOR_RGB2BGR))

        self.dust_map_path = path
        self.lbl_save_status.config(text=f"✓ {os.path.basename(path)}")

    def _write_config(self):
        if self.dust_map_path is None:
            messagebox.showerror("Error", "Save the .tif file first.")
            return

        config_path = filedialog.asksaveasfilename(
            title="Save to config.json",
            initialfile="config.json",
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
        )
        if not config_path:
            return

        with open(config_path, "r", encoding="utf-8") as f:
            content = f.read().strip()
            config = json.loads(content) if content else {}

        config["dust_map"] = {
            "path": self.dust_map_path,
            "objective": self.entry_objective.get(),
            "iso": int(self.spin_iso.get()),
            "exposure_flat": self.entry_expo_flat.get(),
            "exposure_production": self.entry_expo_prod.get(),
            "date": date.today().isoformat(),
        }

        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)

        self.lbl_save_status.config(text="✓ config.json updated")


if __name__ == "__main__":
    root = tk.Tk()
    app = DustMapProfiler(root)
    root.mainloop()
