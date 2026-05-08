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

VALID_IMG_EXTS = {'.rw2', '.cr3', '.nef', '.dng', '.arw', '.jpg', '.jpeg'}
PATCH_HALF = 250  # central 500×500 px patch


class NoiseProfiler:
    def __init__(self, root):
        self.root = root
        self.root.title("Phase 0.5: Noise Profiler — Photon Noise Calibration")
        self.root.minsize(900, 550)

        self.center_patches = None   # [N, 500, 500] float32, green channel
        self.mean_val = None
        self.variance_val = None
        self.photon_scale = None
        self._flat_files = []
        self.tk_img = None

        self.display_mode = tk.StringVar(value="histogram")

        self._setup_ui()
        self.canvas.bind("<Configure>", lambda _: self._refresh_canvas())

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------

    def _setup_ui(self):
        ctrl = ttk.Frame(self.root, padding=10)
        ctrl.pack(side=tk.LEFT, fill=tk.Y)

        ttk.Label(ctrl, text="Phase 0.5 — Noise Profiler",
                  font=("Arial", 11, "bold")).pack(pady=(0, 10))

        # --- 1. Flat frames ---
        acq_frame = ttk.LabelFrame(ctrl, text="1. Flat Frames (ISO 100)", padding=8)
        acq_frame.pack(fill=tk.X, pady=6)

        ttk.Button(acq_frame, text="Select flat frames folder (RAW / JPG)",
                   command=self._load_flat_frames).pack(fill=tk.X, pady=2)

        self.lbl_frames = ttk.Label(acq_frame, text="Flat frames: 0")
        self.lbl_frames.pack(anchor="w", pady=(2, 6))

        ttk.Button(acq_frame, text="Compute photon_scale",
                   command=self._calculate).pack(fill=tk.X, pady=2)

        ttk.Label(acq_frame,
                  text="ISO 100 mandatory — shot noise dominant.\nUniform illuminated background, stable exposure.",
                  font=("Arial", 8), foreground="gray").pack(anchor="w", pady=(6, 0))

        # --- 2. Results ---
        res_frame = ttk.LabelFrame(ctrl, text="2. Results", padding=8)
        res_frame.pack(fill=tk.X, pady=6)

        self.lbl_mean = ttk.Label(res_frame, text="Mean intensity: --")
        self.lbl_mean.pack(anchor="w")
        self.lbl_var = ttk.Label(res_frame, text="Noise variance: --")
        self.lbl_var.pack(anchor="w")

        ttk.Separator(res_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=6)

        self.lbl_scale = ttk.Label(res_frame, text="photon_scale : --",
                                   font=("Arial", 13, "bold"))
        self.lbl_scale.pack(anchor="w")

        # --- 3. Save ---
        save_frame = ttk.LabelFrame(ctrl, text="3. Save", padding=8)
        save_frame.pack(fill=tk.X, pady=6)

        ttk.Button(save_frame, text="⚙️ Save to config.json",
                   command=self._write_config).pack(fill=tk.X, pady=2)

        self.lbl_save_status = ttk.Label(save_frame, text="")
        self.lbl_save_status.pack(anchor="w", pady=(4, 0))

        # --- Display ---
        ttk.Separator(ctrl, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        ttk.Label(ctrl, text="Display:").pack(anchor="w")

        for text, value in [
            ("Noise distribution", "histogram"),
            ("Spatial variance map", "variance_map"),
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
            return img / 255.0
        if rawpy is None:
            messagebox.showerror("Error", "Install rawpy: pip install rawpy")
            return None
        with rawpy.imread(path) as raw:
            rgb16 = raw.postprocess(
                gamma=(1, 1),
                no_auto_bright=True,
                use_camera_wb=True,
                output_bps=16,
            )
        return rgb16.astype(np.float32) / 65535.0

    def _list_raw_files(self, dir_path):
        return sorted([
            os.path.join(dir_path, f)
            for f in os.listdir(dir_path)
            if os.path.splitext(f)[1].lower() in VALID_IMG_EXTS
        ])

    # ------------------------------------------------------------------
    # Actions
    # ------------------------------------------------------------------

    def _load_flat_frames(self):
        dir_path = filedialog.askdirectory(title="Flat frames folder RAW / JPG (ISO 100)")
        if not dir_path:
            return
        self._flat_files = self._list_raw_files(dir_path)
        self.lbl_frames.config(text=f"Flat frames: {len(self._flat_files)}")
        if len(self._flat_files) < 3:
            messagebox.showwarning("Warning",
                                   f"{len(self._flat_files)} file(s) found. "
                                   "5–10 frames recommended.")

    def _calculate(self):
        if len(self._flat_files) < 3:
            messagebox.showerror("Error", "Load the folder first (min. 3 frames).")
            return

        patches = []
        for i, f in enumerate(self._flat_files):
            self.lbl_mean.config(text=f"Loading {i + 1}/{len(self._flat_files)}…")
            self.root.update()
            img = self._load_raw_linear(f)
            if img is None:
                continue
            H, W = img.shape[:2]
            cy, cx = H // 2, W // 2
            patch = img[cy - PATCH_HALF:cy + PATCH_HALF, cx - PATCH_HALF:cx + PATCH_HALF, 1]
            patches.append(patch)

        if len(patches) < 3:
            messagebox.showerror("Error", "Not enough readable RAW frames.")
            return

        self.center_patches = np.array(patches)
        self.mean_val = float(np.mean(self.center_patches))
        self.variance_val = float(np.mean(np.var(self.center_patches, axis=0)))

        if self.variance_val <= 0:
            messagebox.showerror("Error", "Zero variance — identical or black images?")
            return

        self.photon_scale = self.mean_val / self.variance_val

        self.lbl_mean.config(text=f"Mean intensity: {self.mean_val:.4f}")
        self.lbl_var.config(text=f"Noise variance: {self.variance_val:.8f}")
        self.lbl_scale.config(text=f"photon_scale : {self.photon_scale:.1f}")

        self._refresh_canvas()

    # ------------------------------------------------------------------
    # Visualizations
    # ------------------------------------------------------------------

    def _render_histogram(self, w, h):
        img = np.full((h, w, 3), 30, dtype=np.uint8)

        data = self.center_patches.flatten()
        p_lo, p_hi = np.percentile(data, [0.5, 99.5])

        bins = 256
        hist, _ = np.histogram(data, bins=bins, range=(p_lo, p_hi))
        hist_norm = hist / hist.max()

        mx, my = 65, 35
        plot_w = w - mx - 25
        plot_h = h - my - 55

        bar_w = max(1, plot_w // bins)
        for i, v in enumerate(hist_norm):
            x = mx + int(i * plot_w / bins)
            bar_h = int(v * plot_h)
            y1 = my + plot_h - bar_h
            y2 = my + plot_h
            cv2.rectangle(img, (x, y1), (x + bar_w - 1, y2), (100, 180, 255), -1)

        # Vertical mean line
        span = p_hi - p_lo + 1e-9
        mean_x = mx + int((self.mean_val - p_lo) / span * plot_w)
        mean_x = int(np.clip(mean_x, mx, mx + plot_w))
        cv2.line(img, (mean_x, my), (mean_x, my + plot_h), (255, 210, 0), 2)
        cv2.putText(img, f"mean={self.mean_val:.4f}",
                    (mean_x + 5, my + 18),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.42, (255, 210, 0), 1)

        # Axes frame
        cv2.rectangle(img, (mx, my), (mx + plot_w, my + plot_h), (180, 180, 180), 1)

        # X-axis labels (3 values)
        for frac, val in [(0, p_lo), (0.5, (p_lo + p_hi) / 2), (1, p_hi)]:
            x_lbl = mx + int(frac * plot_w)
            cv2.putText(img, f"{val:.3f}", (x_lbl - 18, my + plot_h + 16),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.38, (160, 160, 160), 1)

        # Title
        cv2.putText(img, "Pixel distribution — green channel (all frames)",
                    (mx, my - 12), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (220, 220, 220), 1)

        # photon_scale large at bottom
        cv2.putText(img, f"photon_scale = {self.photon_scale:.1f}",
                    (mx, my + plot_h + 38),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.75, (100, 255, 150), 2)

        return cv2.cvtColor(img, cv2.COLOR_BGR2RGB)

    def _render_variance_map(self, w, h):
        var_map = np.var(self.center_patches, axis=0).astype(np.float32)

        p1, p99 = np.percentile(var_map, [1, 99])
        stretched = ((var_map - p1) / (p99 - p1 + 1e-9) * 255).clip(0, 255).astype(np.uint8)
        colored = cv2.applyColorMap(stretched, cv2.COLORMAP_INFERNO)

        cv2.putText(colored, "Spatial variance map — green channel",
                     (10, 26), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 255, 255), 1)
        cv2.putText(colored, "Bright areas = higher variance (strong noise)",
                     (10, 50), cv2.FONT_HERSHEY_SIMPLEX, 0.44, (200, 200, 200), 1)

        return cv2.cvtColor(colored, cv2.COLOR_BGR2RGB)

    def _refresh_canvas(self):
        if self.center_patches is None:
            return

        cw = self.canvas.winfo_width() or 800
        ch = self.canvas.winfo_height() or 600

        mode = self.display_mode.get()
        if mode == "histogram":
            arr = self._render_histogram(cw, ch)
            pil_img = Image.fromarray(arr)
        else:
            arr = self._render_variance_map(cw, ch)
            pil_img = Image.fromarray(arr)
            scale = min(cw / pil_img.width, ch / pil_img.height, 1.0)
            new_w = max(1, int(pil_img.width * scale))
            new_h = max(1, int(pil_img.height * scale))
            pil_img = pil_img.resize((new_w, new_h), Image.Resampling.LANCZOS)

        self.tk_img = ImageTk.PhotoImage(pil_img)
        self.canvas.delete("all")
        self.canvas.create_image(cw // 2, ch // 2, anchor=tk.CENTER, image=self.tk_img)

    # ------------------------------------------------------------------
    # Export config.json
    # ------------------------------------------------------------------

    def _write_config(self):
        if self.photon_scale is None:
            messagebox.showerror("Error", "Compute photon_scale first.")
            return

        config_path = filedialog.asksaveasfilename(
            title="Save to config.json",
            initialfile="config.json",
            defaultextension=".json",
            filetypes=[("JSON", "*.json")],
        )
        if not config_path:
            return

        try:
            with open(config_path, "r", encoding="utf-8") as f:
                content = f.read().strip()
            config = json.loads(content) if content else {}
        except (FileNotFoundError, json.JSONDecodeError):
            config = {}

        config["poisson_calibration"] = {
            "photon_scale": round(self.photon_scale, 2),
            "mean": round(self.mean_val, 6),
            "variance": round(self.variance_val, 10),
            "n_frames": len(self._flat_files),
            "date": date.today().isoformat(),
        }

        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)

        self.lbl_save_status.config(text="✓ config.json updated")


if __name__ == "__main__":
    root = tk.Tk()
    app = NoiseProfiler(root)
    root.mainloop()