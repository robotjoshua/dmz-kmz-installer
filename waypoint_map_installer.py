"""
Waypoint Map KMZ Installer
Replaces waypoint missions on DJI RC 2 / RC Pro controllers via MTP.

Uses MediaDevices.dll (same library as the original Avenian app) for reliable
MTP file operations via pythonnet.
"""

import io
import os
import re
import shutil
import sys
import tempfile
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

# ── .NET / MediaDevices setup ──────────────────────────────────────────────
import clr

# When frozen by PyInstaller, files are in sys._MEIPASS; otherwise next to script
if getattr(sys, 'frozen', False):
    SCRIPT_DIR = os.path.dirname(sys.executable)
    BUNDLE_DIR = sys._MEIPASS
else:
    SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
    BUNDLE_DIR = SCRIPT_DIR

for candidate in [BUNDLE_DIR, SCRIPT_DIR]:
    dll = os.path.join(candidate, "MediaDevices.dll")
    if os.path.exists(dll):
        sys.path.insert(0, os.path.abspath(candidate))
        break

clr.AddReference("MediaDevices")
from MediaDevices import MediaDevice  # noqa: E402

APP_DIR = BUNDLE_DIR
WAYPOINT_REL = r"Internal shared storage\Android\data\dji.go.v5\files\waypoint"

# DJI drone enum values
DJI_DRONES = {
    "60": "Mini 3 Pro", "67": "Mavic 3 Enterprise", "68": "Mavic 3",
    "77": "Air 3", "89": "Mini 4 Pro", "91": "Air 3S",
    "99": "Mavic 3 Pro", "100": "Matrice 350",
}

SKIP_FOLDERS = {"capability", "map_preview"}


# ── MTP helpers (MediaDevices.dll) ──────────────────────────────────────────

def get_mtp_devices():
    """Return list of connected MTP devices as dicts."""
    result = []
    for dev in MediaDevice.GetDevices():
        result.append({
            "friendly_name": dev.FriendlyName,
            "description": dev.Description,
            "device_id": dev.DeviceId,
        })
    return result


def _with_device(device_id, func):
    """Connect to a device by ID, run func(device), disconnect. Returns func result."""
    for dev in MediaDevice.GetDevices():
        if dev.DeviceId == device_id:
            dev.Connect()
            try:
                return func(dev)
            finally:
                dev.Disconnect()
    raise RuntimeError("Device not found")


def _wp_base(dev):
    """Return the waypoint base path on the device, e.g. '\\Internal shared storage\\Android\\...'"""
    root = "\\"
    for d in dev.GetDirectories(root):
        candidate = d + "\\" + "Android\\data\\dji.go.v5\\files\\waypoint"
        if dev.DirectoryExists(candidate):
            return candidate
    # Fallback: try the known path directly
    known = "\\" + WAYPOINT_REL
    if dev.DirectoryExists(known):
        return known
    raise RuntimeError("Could not find waypoint directory on device")


def list_missions(device_id):
    """List mission folders on device. Returns list of dicts."""
    def _work(dev):
        wp = _wp_base(dev)
        missions = []
        for d in dev.GetDirectories(wp):
            name = d.rsplit("\\", 1)[-1]
            if name in SKIP_FOLDERS:
                continue
            has_file = False
            mission_file = ""
            try:
                files = dev.GetFiles(d)
                for f in files:
                    fn = f.rsplit("\\", 1)[-1]
                    if fn.startswith("."):
                        continue
                    has_file = True
                    mission_file = fn
            except Exception:
                pass
            missions.append({
                "folder_name": name,
                "has_kmz": has_file,
                "mission_file": mission_file,
            })
        return missions
    return _with_device(device_id, _work)


def replace_mission(device_id, mission_folder, src_kmz_path,
                    target_drone_code=None, progress_callback=None):
    """Full mission replacement — mirrors Avenian app's ReplaceMissionSameId:
    1. Normalize/patch KMZ
    2. Delete old files, upload new KMZ
    3. Upload preview images
    4. Stamp folders
    """
    log = []

    def _log(msg):
        log.append(msg)
        if progress_callback:
            progress_callback(msg)

    # ── Normalize KMZ ──
    _log("Normalizing KMZ (flatten + drone code patch)...")
    normalized = normalize_kmz(src_kmz_path, target_drone_code=target_drone_code)

    def _work(dev):
        wp = _wp_base(dev)
        mission_path = wp + "\\" + mission_folder

        if not dev.DirectoryExists(mission_path):
            return False, f"Mission folder not found: {mission_folder}"

        # ── Delete existing files (except image/ folder) ──
        _log("Deleting old mission files...")
        try:
            for f in dev.GetFiles(mission_path):
                fn = f.rsplit("\\", 1)[-1]
                _log(f"  Deleting: {fn}")
                dev.DeleteFile(f)
        except Exception as e:
            _log(f"  Delete warning: {e}")

        # ── Upload normalized KMZ ──
        # MTP stores files with .kmz extension (content-type detection)
        _log("Uploading KMZ file...")
        dest_file = mission_path + "\\" + mission_folder + ".kmz"
        with open(normalized, "rb") as fh:
            from System.IO import MemoryStream
            data = fh.read()
            ms = MemoryStream(data)
            dev.UploadFile(ms, dest_file)
            ms.Close()
        _log(f"  Uploaded {len(data)} bytes as {mission_folder}.kmz")

        # Verify upload
        try:
            files_after = list(dev.GetFiles(mission_path))
            _log(f"  Files in mission folder: {len(files_after)}")
        except Exception:
            pass

        # ── Generate and upload preview image ──
        _log("Generating preview image...")
        preview_data = _make_preview_image("Mission Updated")

        # Upload to map_preview/ root
        map_preview_path = wp + "\\map_preview"
        preview_file = map_preview_path + "\\" + mission_folder + ".jpg"
        _log(f"Uploading preview to map_preview/...")
        try:
            if dev.FileExists(preview_file):
                dev.DeleteFile(preview_file)
            ms = MemoryStream(preview_data)
            dev.UploadFile(ms, preview_file)
            ms.Close()
        except Exception as e:
            _log(f"  Preview upload warning: {e}")

        # Upload to map_preview/<missionId>/ subfolder
        mp_sub = map_preview_path + "\\" + mission_folder
        if dev.DirectoryExists(mp_sub):
            mp_sub_file = mp_sub + "\\" + mission_folder + ".jpg"
            _log("Uploading preview to map_preview subfolder...")
            try:
                if dev.FileExists(mp_sub_file):
                    dev.DeleteFile(mp_sub_file)
                ms = MemoryStream(preview_data)
                dev.UploadFile(ms, mp_sub_file)
                ms.Close()
            except Exception as e:
                _log(f"  Subfolder preview warning: {e}")

        # Upload to mission/image/ as 'cover'
        image_path = mission_path + "\\image"
        if dev.DirectoryExists(image_path):
            cover_file = image_path + "\\cover"
            _log("Uploading cover to mission/image/...")
            try:
                if dev.FileExists(cover_file):
                    dev.DeleteFile(cover_file)
                ms = MemoryStream(preview_data)
                dev.UploadFile(ms, cover_file)
                ms.Close()
            except Exception as e:
                _log(f"  Cover upload warning: {e}")

        # ── Clean up any leftover .stamp files ──
        _log("Cleaning up junk files...")
        for check_path in [mission_path, map_preview_path]:
            try:
                for f in dev.GetFiles(check_path):
                    fn = f.rsplit("\\", 1)[-1]
                    if fn == ".stamp":
                        dev.DeleteFile(f)
                        _log(f"  Removed .stamp from {check_path}")
            except Exception:
                pass

        _log("Done!")
        return True, "\n".join(log)

    try:
        return _with_device(device_id, _work)
    except Exception as e:
        return False, f"Error: {e}\n\nLog:\n" + "\n".join(log)


def verify_mission_on_device(device_id, mission_folder):
    """Pull mission file from device and inspect it. Returns report string."""
    def _work(dev):
        wp = _wp_base(dev)
        mission_path = wp + "\\" + mission_folder
        report = [f"Mission: {mission_folder}\n"]

        # List all files
        try:
            files = list(dev.GetFiles(mission_path))
            report.append(f"Files in folder: {len(files)}")
            for f in files:
                fn = f.rsplit("\\", 1)[-1]
                report.append(f"  {fn}")
        except Exception as e:
            report.append(f"Error listing: {e}")
            return "\n".join(report)

        # Find the mission file (not .stamp, not in subfolders)
        mission_file = None
        for f in files:
            fn = f.rsplit("\\", 1)[-1]
            if not fn.startswith("."):
                mission_file = f
                break

        if not mission_file:
            report.append("\nNo mission file found!")
            return "\n".join(report)

        # Download and inspect
        from System.IO import MemoryStream
        ms = MemoryStream()
        dev.DownloadFile(mission_file, ms)
        data = bytes(ms.ToArray())
        ms.Close()
        report.append(f"\nDownloaded: {len(data)} bytes")

        try:
            bio = io.BytesIO(data)
            with zipfile.ZipFile(bio) as zf:
                names = zf.namelist()
                report.append(f"ZIP contents: {names}")
                for n in names:
                    if n.lower().endswith(".kml"):
                        kml = zf.read(n).decode("utf-8", errors="replace")
                        m = re.search(r"<wpml:droneEnumValue>(\d+)</wpml:droneEnumValue>", kml)
                        if m:
                            code = m.group(1)
                            drone = DJI_DRONES.get(code, f"Unknown ({code})")
                            report.append(f"\ndroneEnumValue: {code} ({drone})")
                    if n.lower().endswith(".wpml"):
                        wpml = zf.read(n).decode("utf-8", errors="replace")
                        wps = len(re.findall(r"<Placemark>", wpml))
                        report.append(f"Waypoints: {wps}")
                has_sub = any("/" in n for n in names if not n.endswith("/"))
                if has_sub:
                    report.append("\nWARNING: Files in subfolders!")
                else:
                    report.append("\nStructure: OK (files at ZIP root)")
        except zipfile.BadZipFile:
            report.append("NOT a valid ZIP/KMZ!")
            report.append(f"Header: {data[:16].hex()}")
        except Exception as e:
            report.append(f"Parse error: {e}")

        return "\n".join(report)

    try:
        return _with_device(device_id, _work)
    except Exception as e:
        return f"Error: {e}"


# ── KMZ helpers ─────────────────────────────────────────────────────────────

def normalize_kmz(src_kmz_path, target_drone_code=None):
    """Repackage a KMZ: flatten wpmz/ subfolder, patch droneEnumValue."""
    temp_dir = os.path.join(tempfile.gettempdir(), "dji_kmz_normalize")
    os.makedirs(temp_dir, exist_ok=True)
    out_path = os.path.join(temp_dir, "normalized.kmz")

    with zipfile.ZipFile(src_kmz_path, "r") as zf_in:
        names = zf_in.namelist()
        needs_flatten = any(
            "/" in n and (n.lower().endswith(".kml") or n.lower().endswith(".wpml"))
            for n in names
        )
        needs_drone_patch = False
        if target_drone_code:
            for n in names:
                if n.lower().endswith(".kml"):
                    data = zf_in.read(n).decode("utf-8", errors="replace")
                    m = re.search(r"<wpml:droneEnumValue>(\d+)</wpml:droneEnumValue>", data)
                    if m and m.group(1) != target_drone_code:
                        needs_drone_patch = True
                    break

        if not needs_flatten and not needs_drone_patch:
            return src_kmz_path

        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf_out:
            for n in names:
                if n.endswith("/"):
                    continue
                flat_name = n.split("/")[-1] if needs_flatten and "/" in n else n
                data = zf_in.read(n)
                if needs_drone_patch and (
                    flat_name.lower().endswith(".kml") or flat_name.lower().endswith(".wpml")
                ):
                    text = data.decode("utf-8", errors="replace")
                    text = re.sub(
                        r"(<wpml:droneEnumValue>)\d+(</wpml:droneEnumValue>)",
                        rf"\g<1>{target_drone_code}\g<2>", text
                    )
                    data = text.encode("utf-8")
                zf_out.writestr(flat_name, data)
    return out_path


def _make_preview_image(text="Waypoint Mission", width=400, height=240):
    """Generate a mission preview JPEG."""
    img = Image.new("RGB", (width, height), color=(53, 53, 53))
    draw = ImageDraw.Draw(img)
    cx, cy, r = 100, height // 2, 60
    draw.ellipse([cx - r, cy - r, cx + r, cy + r], fill=(0, 192, 216))
    for i in range(5):
        dx = cx - 30 + i * 15
        dy = cy - 30 + (i % 3) * 20 - 10
        draw.ellipse([dx - 4, dy - 4, dx + 4, dy + 4], fill="white")
    try:
        font = ImageFont.truetype("segoeui.ttf", 28)
    except Exception:
        font = ImageFont.load_default()
    draw.text((180, height // 2 - 15), text, fill="white", font=font)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=80)
    return buf.getvalue()


def parse_kmz_info(kmz_path):
    """Extract info from a KMZ file."""
    info = {
        "name": Path(kmz_path).stem, "description": "", "placemarks": 0,
        "overlays": 0, "format": "unknown", "drone": "", "speed": "",
    }
    try:
        with zipfile.ZipFile(kmz_path, "r") as zf:
            names = zf.namelist()
            has_wpml = any(n.lower().endswith(".wpml") for n in names)
            if has_wpml:
                info["format"] = "DJI Waypoint"
                wpml_file = [n for n in names if n.lower().endswith(".wpml")][0]
                with zf.open(wpml_file) as wf:
                    tree = ET.parse(wf)
                    root = tree.getroot()
                    info["placemarks"] = len(root.findall(
                        ".//{http://www.opengis.net/kml/2.2}Placemark"))
                    if info["placemarks"] == 0:
                        info["placemarks"] = len(root.findall(".//Placemark"))
                kml_files = [n for n in names if n.lower().endswith(".kml")]
                if kml_files:
                    with zf.open(kml_files[0]) as kf:
                        tree = ET.parse(kf)
                        root = tree.getroot()
                        ns = "http://www.dji.com/wpmz/1.0.2"
                        de = root.find(f".//{{{ns}}}droneEnumValue")
                        if de is not None and de.text:
                            info["drone"] = DJI_DRONES.get(de.text, f"DJI ({de.text})")
                        sp = root.find(f".//{{{ns}}}globalTransitionalSpeed")
                        if sp is not None and sp.text:
                            info["speed"] = f"{sp.text} m/s"
                return info
            kml_files = [f for f in names if f.lower().endswith(".kml")]
            if kml_files:
                info["format"] = "KML"
                with zf.open(kml_files[0]) as kml:
                    tree = ET.parse(kml)
                    root = tree.getroot()
                    kns = ""
                    if root.tag.startswith("{"):
                        kns = root.tag.split("}")[0] + "}"
                    info["placemarks"] = len(root.findall(f".//{kns}Placemark"))
    except Exception:
        pass
    return info


def format_size(size_bytes):
    for unit in ("B", "KB", "MB", "GB"):
        if size_bytes < 1024:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024
    return f"{size_bytes:.1f} TB"


# ── Main Application ─────────────────────────────────────────────────────────

# Color palette  —  DPM target theme
C_BG       = "#3A3A3A"        # dark charcoal background
C_SURFACE  = "#444444"        # card surfaces
C_SURFACE2 = "#4F4F4F"        # lighter surface (hover)
C_BORDER   = "#F5C800"        # yellow border (logo ring)
C_TEXT     = "#F0ECD5"        # cream primary text (logo light quadrants)
C_MUTED    = "#B0AB9E"        # muted / secondary text
C_ACCENT   = "#F5C800"        # yellow accent
C_ACCENT2  = "#C49A00"        # darker yellow
C_GREEN    = "#F5C800"        # success yellow
C_WARN     = "#F5C800"        # warning yellow
C_DANGER   = "#ef476f"        # error/mismatch red
C_STEP     = "#F5C800"        # step number yellow


class WaypointMapInstaller(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Drone Pilot KMZ Installer")
        self.geometry("820x740")
        self.minsize(720, 640)
        self.configure(bg=C_BG)

        icon_path = os.path.join(APP_DIR, "assets", "logo.png")
        if os.path.exists(icon_path):
            try:
                _icon = Image.open(icon_path).resize((32, 32), Image.LANCZOS)
                from PIL import ImageTk
                self._icon_img = ImageTk.PhotoImage(_icon)
                self.iconphoto(True, self._icon_img)
            except Exception:
                pass

        self.kmz_file = None
        self._devices = []
        self._dest_info = None
        self._missions = []

        self._build_styles()
        self._build_ui()
        self.refresh_devices()

    def _build_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure(".", background=C_BG, foreground=C_TEXT, font=("Segoe UI", 10))

        # Header
        style.configure("Header.TLabel", font=("Segoe UI", 20, "bold"),
                         foreground=C_TEXT, background=C_BG)
        style.configure("Sub.TLabel", font=("Segoe UI", 10),
                         foreground=C_MUTED, background=C_BG)

        # Step labels
        style.configure("Step.TLabel", font=("Segoe UI", 10, "bold"),
                         foreground=C_ACCENT, background=C_BG)

        # Cards / LabelFrames
        style.configure("Card.TFrame", background=C_SURFACE)
        style.configure("Section.TLabelframe", background=C_SURFACE,
                         relief="flat", borderwidth=0)
        style.configure("Section.TLabelframe.Label", font=("Segoe UI", 10, "bold"),
                         foreground=C_MUTED, background=C_SURFACE)

        # Labels inside cards
        style.configure("Info.TLabel", foreground=C_MUTED, background=C_SURFACE,
                         font=("Segoe UI", 9))
        style.configure("FileInfo.TLabel", foreground=C_TEXT, background=C_SURFACE,
                         font=("Segoe UI Semibold", 9))
        style.configure("Warn.TLabel", foreground=C_WARN, background=C_BG,
                         font=("Segoe UI", 9))
        style.configure("CardLabel.TLabel", foreground=C_TEXT, background=C_SURFACE,
                         font=("Segoe UI", 10))

        # Accent button (install)
        style.configure("Accent.TButton", font=("Segoe UI", 11, "bold"),
                         foreground="#1A1A1A", background="#F5C800",
                         borderwidth=0, padding=(20, 10))
        style.map("Accent.TButton",
                   background=[("active", "#C49A00"), ("disabled", "#5A5A5A")])

        # Secondary button
        style.configure("Secondary.TButton", font=("Segoe UI", 9),
                         foreground=C_TEXT, background=C_SURFACE2,
                         borderwidth=0, padding=(14, 6))
        style.map("Secondary.TButton",
                   background=[("active", C_BORDER), ("disabled", C_SURFACE)])

        # Progress bar
        style.configure("Custom.Horizontal.TProgressbar",
                         troughcolor=C_SURFACE2, background=C_GREEN, thickness=6)

        # Treeview
        style.configure("Treeview", background=C_SURFACE, foreground=C_TEXT,
                         fieldbackground=C_SURFACE, borderwidth=0,
                         font=("Consolas", 9), rowheight=28)
        style.configure("Treeview.Heading", background=C_SURFACE2,
                         foreground=C_MUTED, font=("Segoe UI", 9, "bold"),
                         borderwidth=0)
        style.map("Treeview",
                   background=[("selected", C_ACCENT2)],
                   foreground=[("selected", "#ffffff")])

        # Combobox
        style.configure("TCombobox", fieldbackground=C_SURFACE2,
                         background=C_SURFACE2, foreground=C_TEXT,
                         borderwidth=0, padding=4)
        style.map("TCombobox", fieldbackground=[("readonly", C_SURFACE2)])

        # Scrollbar
        style.configure("TScrollbar", background=C_SURFACE2,
                         troughcolor=C_SURFACE, borderwidth=0)

    def _make_step_label(self, parent, number, text):
        """Create a styled step indicator: circled number + text."""
        frame = tk.Frame(parent, bg=C_BG)
        # Step number badge
        badge = tk.Canvas(frame, width=26, height=26, bg=C_BG,
                          highlightthickness=0)
        badge.create_oval(2, 2, 24, 24, fill=C_STEP, outline="")
        badge.create_text(13, 13, text=str(number), fill="white",
                          font=("Segoe UI", 9, "bold"))
        badge.pack(side="left", padx=(0, 8))
        tk.Label(frame, text=text, bg=C_BG, fg=C_ACCENT,
                 font=("Segoe UI", 10, "bold")).pack(side="left")
        return frame

    def _make_card(self, parent, **pack_kw):
        """Create a rounded-feeling card frame."""
        outer = tk.Frame(parent, bg=C_BORDER, padx=1, pady=1)
        outer.pack(fill="x", padx=24, pady=(0, 6), **pack_kw)
        inner = tk.Frame(outer, bg=C_SURFACE, padx=14, pady=12)
        inner.pack(fill="both", expand=True)
        return inner

    def _build_ui(self):
        # ── Scrollable container ──
        container = tk.Frame(self, bg=C_BG)
        container.pack(fill="both", expand=True)

        # ── Header ──
        header = tk.Frame(container, bg=C_BG)
        header.pack(fill="x", padx=24, pady=(18, 4))

        logo_path = os.path.join(APP_DIR, "assets", "logo.png")
        if os.path.exists(logo_path):
            try:
                # Load via PIL for proper resizing, then convert to PhotoImage
                _pil_logo = Image.open(logo_path).resize((48, 48), Image.LANCZOS)
                from PIL import ImageTk
                self._logo_img = ImageTk.PhotoImage(_pil_logo)
                tk.Label(header, image=self._logo_img, bg=C_BG).pack(
                    side="left", padx=(0, 14))
            except Exception:
                pass

        tf = tk.Frame(header, bg=C_BG)
        tf.pack(side="left")
        tk.Label(tf, text="Drone Pilot KMZ Installer", bg=C_BG,
                 fg=C_TEXT, font=("Segoe UI", 20, "bold")).pack(anchor="w")
        tk.Label(tf, text="Transfer KMZ waypoint missions to DJI RC 2 / RC Pro",
                 bg=C_BG, fg=C_MUTED, font=("Segoe UI", 10)).pack(anchor="w")

        # Thin accent line (yellow)
        tk.Frame(container, bg="#F5C800", height=2).pack(fill="x", padx=24, pady=(6, 12))

        # ── Step 1: Device ──
        self._make_step_label(container, 1, "Select your controller").pack(
            anchor="w", padx=24, pady=(0, 4))
        card1 = self._make_card(container)

        dev_row = tk.Frame(card1, bg=C_SURFACE)
        dev_row.pack(fill="x")
        self.device_var = tk.StringVar()
        self.device_combo = ttk.Combobox(dev_row, textvariable=self.device_var,
                                          state="readonly", width=48)
        self.device_combo.pack(side="left", padx=(0, 8))
        self.device_combo.bind("<<ComboboxSelected>>", self._on_device_selected)
        ttk.Button(dev_row, text="Refresh", style="Secondary.TButton",
                    command=self.refresh_devices).pack(side="left")

        # ── Step 2: KMZ File ──
        self._make_step_label(container, 2, "Select the .KMZ file to install").pack(
            anchor="w", padx=24, pady=(8, 4))
        card2 = self._make_card(container)

        kmz_row = tk.Frame(card2, bg=C_SURFACE)
        kmz_row.pack(fill="x")
        ttk.Button(kmz_row, text="Browse KMZ...", style="Secondary.TButton",
                    command=self.browse_kmz).pack(side="left", padx=(0, 10))
        self.kmz_label = tk.Label(kmz_row, text="No file selected", bg=C_SURFACE,
                                   fg=C_MUTED, font=("Segoe UI", 10))
        self.kmz_label.pack(side="left", fill="x")

        self.kmz_info_label = tk.Label(card2, text="", bg=C_SURFACE, fg=C_GREEN,
                                        font=("Segoe UI", 9))
        self.kmz_info_label.pack(anchor="w", pady=(6, 0))

        # Drone target selector
        drone_row = tk.Frame(card2, bg=C_SURFACE)
        drone_row.pack(fill="x", pady=(6, 0))
        tk.Label(drone_row, text="Target drone:", bg=C_SURFACE, fg=C_MUTED,
                 font=("Segoe UI", 9)).pack(side="left", padx=(0, 8))
        self.drone_var = tk.StringVar(value="89 — Mini 4 Pro")
        drone_options = [f"{k} — {v}" for k, v in
                         sorted(DJI_DRONES.items(), key=lambda x: x[1])]
        self.drone_combo = ttk.Combobox(drone_row, textvariable=self.drone_var,
                                         values=drone_options, state="readonly", width=25)
        self.drone_combo.pack(side="left")
        self.drone_mismatch_label = tk.Label(
            drone_row, text="", bg=C_SURFACE, fg=C_DANGER,
            font=("Segoe UI", 9, "bold"))
        self.drone_mismatch_label.pack(side="left", padx=(10, 0))

        # ── Step 3: Missions ──
        self._make_step_label(container, 3, "Select a mission to replace").pack(
            anchor="w", padx=24, pady=(8, 4))

        # Mission card (expandable)
        mission_outer = tk.Frame(container, bg=C_BORDER, padx=1, pady=1)
        mission_outer.pack(fill="both", expand=True, padx=24, pady=(0, 6))
        mission_card = tk.Frame(mission_outer, bg=C_SURFACE, padx=14, pady=10)
        mission_card.pack(fill="both", expand=True)

        mbtn_row = tk.Frame(mission_card, bg=C_SURFACE)
        mbtn_row.pack(fill="x", pady=(0, 8))
        ttk.Button(mbtn_row, text="List Missions", style="Secondary.TButton",
                    command=self.do_list_missions).pack(side="left")
        ttk.Button(mbtn_row, text="Verify Selected", style="Secondary.TButton",
                    command=self.do_verify).pack(side="left", padx=(8, 0))
        tk.Label(mbtn_row, text="Replaces the selected mission's KMZ on device",
                 bg=C_SURFACE, fg=C_MUTED, font=("Segoe UI", 8)).pack(
                     side="right", padx=(0, 4))

        tree_frame = tk.Frame(mission_card, bg=C_SURFACE)
        tree_frame.pack(fill="both", expand=True)
        cols = ("name", "has_file")
        self.mission_tree = ttk.Treeview(tree_frame, columns=cols,
                                          show="headings", height=5,
                                          selectmode="browse")
        self.mission_tree.heading("name", text="Mission ID")
        self.mission_tree.heading("has_file", text="Has KMZ")
        self.mission_tree.column("name", width=460)
        self.mission_tree.column("has_file", width=80, anchor="center")
        sb = ttk.Scrollbar(tree_frame, orient="vertical",
                            command=self.mission_tree.yview)
        self.mission_tree.configure(yscrollcommand=sb.set)
        self.mission_tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # ── Bottom bar: progress + install ──
        bottom = tk.Frame(container, bg=C_BG)
        bottom.pack(fill="x", padx=24, pady=(8, 16))

        self.progress = ttk.Progressbar(bottom, mode="determinate",
                                         style="Custom.Horizontal.TProgressbar")
        self.progress.pack(fill="x", pady=(0, 10))

        status_row = tk.Frame(bottom, bg=C_BG)
        status_row.pack(fill="x")
        self.status_var = tk.StringVar(value="Ready  \u2014  plug in your DJI RC via USB")
        self.status_label = tk.Label(status_row, textvariable=self.status_var,
                                      bg=C_BG, fg=C_MUTED, font=("Segoe UI", 9),
                                      anchor="w")
        self.status_label.pack(side="left", fill="x", expand=True)
        self.install_btn = ttk.Button(status_row, text="   Install to RC   ",
                                       style="Accent.TButton",
                                       command=self.start_install)
        self.install_btn.pack(side="right")

    # ── Device ──────────────────────────────────────────────────────────────

    def refresh_devices(self):
        self.status_var.set("Scanning for devices...")
        self.update_idletasks()
        self._devices = get_mtp_devices()
        display = [f"{d['friendly_name']}" for d in self._devices]
        self.device_combo["values"] = display if display else [
            "No devices found \u2014 check USB connection"]
        if display:
            self.device_combo.current(0)
            self._on_device_selected(None)
        else:
            self.device_var.set("No devices found \u2014 check USB connection")
            self._dest_info = None
        self.status_var.set(
            f"\u2713  Found {len(self._devices)} device(s)" if self._devices
            else "\u2717  No device detected")

    def _on_device_selected(self, _event):
        idx = self.device_combo.current()
        if 0 <= idx < len(self._devices):
            self._dest_info = self._devices[idx]
            self.status_var.set(f"\u2713  Connected: {self._dest_info['friendly_name']}")

    # ── KMZ ─────────────────────────────────────────────────────────────────

    def browse_kmz(self):
        path = filedialog.askopenfilename(
            title="Select KMZ File",
            filetypes=[("DJI KMZ files", "*.kmz"), ("All Files", "*.*")],
            initialdir=os.path.expanduser("~/Downloads"),
        )
        if path:
            self.kmz_file = path
            info = parse_kmz_info(path)
            size = format_size(os.path.getsize(path))
            self.kmz_label.config(text=f"\u2713  {os.path.basename(path)}", fg=C_TEXT)
            details = f"{info['format']}  \u00b7  {info['placemarks']} waypoints  \u00b7  {size}"
            if info.get("speed"):
                details += f"  \u00b7  {info['speed']}"
            self.kmz_info_label.config(text=details)

            # Check drone mismatch
            target_code = self.drone_var.get().split(" — ")[0]
            src_code = self._get_kmz_drone_code(path)
            if src_code and src_code != target_code:
                self.drone_mismatch_label.config(
                    text="\u2713  Will auto-patch to Mini 4 Pro", fg=C_GREEN)
            else:
                self.drone_mismatch_label.config(text="", fg=C_GREEN)

    def _get_kmz_drone_code(self, path):
        try:
            with zipfile.ZipFile(path) as zf:
                for n in zf.namelist():
                    if n.lower().endswith(".kml"):
                        data = zf.read(n).decode("utf-8", errors="replace")
                        m = re.search(
                            r"<wpml:droneEnumValue>(\d+)</wpml:droneEnumValue>", data)
                        if m:
                            return m.group(1)
        except Exception:
            pass
        return None

    # ── Missions ────────────────────────────────────────────────────────────

    def do_list_missions(self):
        if not self._dest_info:
            messagebox.showwarning("No Device", "Select a device first.")
            return
        self.status_var.set("Scanning missions...")
        self.update_idletasks()
        self.mission_tree.delete(*self.mission_tree.get_children())

        try:
            self._missions = list_missions(self._dest_info["device_id"])
        except Exception as e:
            self.status_var.set(f"Error: {e}")
            messagebox.showerror("Error", str(e))
            return

        if not self._missions:
            self.status_var.set("No missions found.")
            messagebox.showinfo(
                "No Missions",
                "No waypoint missions on device.\n\n"
                "Create at least one mission in DJI Fly first.")
            return

        for m in self._missions:
            self.mission_tree.insert("", "end", iid=m["folder_name"], values=(
                m["folder_name"],
                "Yes" if m["has_kmz"] else "No",
            ))
        self.status_var.set(f"\u2713  Found {len(self._missions)} mission(s)")

    # ── Verify ──────────────────────────────────────────────────────────────

    def do_verify(self):
        if not self._dest_info:
            messagebox.showwarning("No Device", "Select a device first.")
            return
        sel = self.mission_tree.selection()
        if not sel:
            messagebox.showwarning("No Selection", "Select a mission to verify.")
            return
        self.status_var.set("Verifying...")
        self.update_idletasks()
        report = verify_mission_on_device(self._dest_info["device_id"], sel[0])
        self.status_var.set("Verify complete")
        messagebox.showinfo("Mission Verification", report)

    # ── Install ─────────────────────────────────────────────────────────────

    def start_install(self):
        if not self._dest_info:
            messagebox.showwarning("No Device", "Select a device first (Step 1).")
            return
        if not self.kmz_file:
            messagebox.showwarning("No KMZ", "Select a .KMZ file first (Step 2).")
            return
        sel = self.mission_tree.selection()
        if not sel:
            messagebox.showwarning("No Mission",
                                    "Select a mission to replace (Step 3).\n\n"
                                    "Click 'List Missions' first if you haven't.")
            return

        mission_folder = sel[0]
        if not messagebox.askyesno(
            "Confirm Replace",
            f"REPLACE mission:\n  {mission_folder}\n\n"
            f"with KMZ:\n  {os.path.basename(self.kmz_file)}\n\nContinue?"
        ):
            return

        self.install_btn.config(state="disabled")
        self.progress["value"] = 0
        self.progress["maximum"] = 100
        thread = threading.Thread(target=self._install_worker,
                                   args=(mission_folder,), daemon=True)
        thread.start()

    def _install_worker(self, mission_folder):
        target_code = self.drone_var.get().split(" — ")[0]
        step = [0]

        def on_progress(msg):
            step[0] += 1
            pct = min(95, step[0] * 12)
            self.after(0, self.status_var.set, msg)
            self.after(0, self._set_progress, pct)

        on_progress("Starting install...")

        ok, detail = replace_mission(
            self._dest_info["device_id"], mission_folder, self.kmz_file,
            target_drone_code=target_code,
            progress_callback=on_progress,
        )

        self.after(0, self._set_progress, 100)
        if ok:
            self.after(0, self._install_done_ok, mission_folder, target_code)
        else:
            self.after(0, self._install_done_err, detail)

    def _set_progress(self, val):
        self.progress["value"] = val

    def _install_done_ok(self, mission_folder, target_code):
        self.install_btn.config(state="normal")
        drone_name = DJI_DRONES.get(target_code, target_code)
        self.status_var.set(f"\u2713  Installed into {mission_folder}  \u2014  ready for next install")

        # Auto-refresh the mission list so Has KMZ column updates
        if self._dest_info:
            try:
                missions = list_missions(self._dest_info["device_id"])
                self._missions = missions
                self.mission_tree.delete(*self.mission_tree.get_children())
                for m in missions:
                    self.mission_tree.insert("", "end", iid=m["folder_name"], values=(
                        m["folder_name"],
                        "Yes" if m["has_kmz"] else "No",
                    ))
            except Exception:
                pass

        # Clear KMZ selection so user is ready to pick the next file
        self.kmz_file = None
        self.kmz_label.config(text="No file selected", fg=C_MUTED)
        self.kmz_info_label.config(text="")
        self.drone_mismatch_label.config(text="")

        messagebox.showinfo(
            "Install Complete",
            f"Installed KMZ into mission:\n"
            f"  {mission_folder}\n"
            f"  Drone: {target_code} ({drone_name})\n\n"
            f"Steps completed:\n"
            f"  \u2713 Old files deleted\n"
            f"  \u2713 Normalized KMZ uploaded\n"
            f"  \u2713 Preview images updated\n"
            f"  \u2713 Junk files cleaned\n\n"
            f"You can install more missions now \u2014 no restart needed between installs.\n"
            f"Restart the DJI RC once when you're done with all installs."
        )

    def _install_done_err(self, err):
        self.install_btn.config(state="normal")
        self.status_var.set("\u2717  Install failed")
        messagebox.showerror("Install Failed", str(err))


if __name__ == "__main__":
    app = WaypointMapInstaller()
    app.mainloop()
