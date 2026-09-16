from __future__ import annotations

import argparse
import csv
import hashlib
import io
import json
import logging
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
import xml.etree.ElementTree as ET
import queue
import threading
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.utils.cell import coordinate_to_tuple
from PIL import Image, ImageEnhance, ImageOps


PROJECT_DIR = Path(__file__).resolve().parent
SUPPORTED_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp"}
MONTHS_ID = {
    1: "JANUARI", 2: "FEBRUARI", 3: "MARET", 4: "APRIL",
    5: "MEI", 6: "JUNI", 7: "JULI", 8: "AGUSTUS",
    9: "SEPTEMBER", 10: "OKTOBER", 11: "NOVEMBER", 12: "DESEMBER",
}
DATE_RANGE_PATTERN = re.compile(
    r"(?P<y1>20\d{2})\s*[/\-.]\s*(?P<m1>\d{1,2})\s*[/\-.]\s*(?P<d1>\d{1,2})"
    r"\s*(?:~|–|—|-|TO)\s*"
    r"(?P<y2>20\d{2})\s*[/\-.]\s*(?P<m2>\d{1,2})\s*[/\-.]\s*(?P<d2>\d{1,2})",
    re.IGNORECASE,
)


@contextmanager
def timed(stage: str):
    started = time.perf_counter()
    logging.info("Mulai: %s", stage)
    try:
        yield
    finally:
        logging.info("Durasi %s: %.2f detik", stage, time.perf_counter() - started)


def file_version(path: Path):
    stat = path.stat()
    return stat.st_size, stat.st_mtime_ns


def load_env_file(path: Path) -> None:
    if not path.exists():
        return
    for raw_line in path.read_text(encoding="utf-8-sig").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def expand_path(value: str, base: Path = PROJECT_DIR) -> Path:
    path = Path(os.path.expandvars(os.path.expanduser(value.strip())))
    return path if path.is_absolute() else (base / path).resolve()


def env_int(name: str, default: int) -> int:
    try:
        return int(os.getenv(name, str(default)).strip())
    except ValueError as exc:
        raise ValueError(f"{name} harus berupa angka.") from exc


def env_bool(name: str, default: bool) -> bool:
    raw = os.getenv(name, str(default)).strip().lower()
    if raw in {"1", "true", "yes", "ya", "on"}:
        return True
    if raw in {"0", "false", "no", "tidak", "off"}:
        return False
    raise ValueError(f"{name} harus true atau false.")


@dataclass(frozen=True)
class Settings:
    screenshot_folder: Path
    drive_report_folder: Path
    backup_folder: Path
    tesseract_command: str
    max_image_width_px: int
    max_image_height_px: int
    crop_speed_summary: bool
    allow_replace_existing: bool
    timezone_offset_hours: int
    notification_mode: str
    n8n_webhook_url: str
    n8n_webhook_token: str
    whatsapp_access_token: str
    whatsapp_phone_number_id: str
    whatsapp_to: str
    whatsapp_api_version: str

    @classmethod
    def from_env(cls) -> "Settings":
        load_env_file(PROJECT_DIR / ".env")
        return cls(
            screenshot_folder=expand_path(os.getenv(
                "SCREENSHOT_FOLDER", r"%USERPROFILE%\OneDrive\Pictures\Screenshots"
            )),
            drive_report_folder=expand_path(os.getenv(
                "DRIVE_REPORT_FOLDER", r"G:\My Drive\Latihan OPD Automation\SEPTEMBER"
            )),
            backup_folder=expand_path(os.getenv("BACKUP_FOLDER", "backups")),
            tesseract_command=os.getenv(
                "TESSERACT_COMMAND", r"C:\Program Files\Tesseract-OCR\tesseract.exe"
            ).strip(),
            max_image_width_px=env_int("MAX_IMAGE_WIDTH_PX", 1220),
            max_image_height_px=env_int("MAX_IMAGE_HEIGHT_PX", 330),
            crop_speed_summary=env_bool("CROP_SPEED_SUMMARY", True),
            allow_replace_existing=env_bool("ALLOW_REPLACE_EXISTING", False),
            timezone_offset_hours=env_int("TIMEZONE_OFFSET_HOURS", 7),
            notification_mode=os.getenv("NOTIFICATION_MODE", "console").strip().lower(),
            n8n_webhook_url=os.getenv("N8N_WEBHOOK_URL", "").strip(),
            n8n_webhook_token=os.getenv("N8N_WEBHOOK_TOKEN", "").strip(),
            whatsapp_access_token=os.getenv("WHATSAPP_ACCESS_TOKEN", "").strip(),
            whatsapp_phone_number_id=os.getenv("WHATSAPP_PHONE_NUMBER_ID", "").strip(),
            whatsapp_to=os.getenv("WHATSAPP_TO", "").strip(),
            whatsapp_api_version=os.getenv("WHATSAPP_API_VERSION", "v23.0").strip() or "v23.0",
        )


class StateStore:
    def __init__(self, path: Path) -> None:
        self.path = path
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.data = self._load()

    def _load(self) -> dict[str, Any]:
        if not self.path.exists():
            return {"processed": {}}
        try:
            return json.loads(self.path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            logging.warning("State tidak terbaca; menggunakan state baru.")
            return {"processed": {}}

    def save(self) -> None:
        temp = self.path.with_suffix(".tmp")
        temp.write_text(json.dumps(self.data, indent=2, ensure_ascii=False), encoding="utf-8")
        os.replace(temp, self.path)

    def was_processed(self, key: str) -> bool:
        return key in self.data.setdefault("processed", {})

    def mark_processed(self, key: str, result: dict[str, Any]) -> None:
        processed = self.data.setdefault("processed", {})
        processed[key] = result
        if len(processed) > 2000:
            for old_key in list(processed)[: len(processed) - 2000]:
                processed.pop(old_key, None)
        self.save()


def local_now(settings: Settings) -> datetime:
    return datetime.now(timezone(timedelta(hours=settings.timezone_offset_hours)))


def setup_logging() -> None:
    logs = PROJECT_DIR / "logs"
    logs.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.FileHandler(logs / f"drive_watcher_{datetime.now():%Y-%m-%d}.log", encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )


def wait_until_stable(path: Path, timeout_seconds: float = 20.0) -> bool:
    deadline, previous, stable = time.monotonic() + timeout_seconds, -1, 0
    while time.monotonic() < deadline:
        try:
            current = path.stat().st_size
        except OSError:
            time.sleep(0.4)
            continue
        stable = stable + 1 if current > 0 and current == previous else 0
        if stable >= 3:
            return True
        previous = current
        time.sleep(0.5)
    return False


def fingerprint(path: Path) -> str:
    stat = path.stat()
    value = f"{path.resolve()}|{stat.st_size}|{stat.st_mtime_ns}".encode()
    return hashlib.sha256(value).hexdigest()


def parse_screenshot_filename(path: Path) -> tuple[str, str]:
    match = re.match(r"^(?P<project>.+?)[_\-\s]+(?P<link>METRO|BROADBAND)$", path.stem, re.I)
    if not match:
        raise ValueError(
            "Nama harus berakhir dengan _METRO atau _BROADBAND, misalnya "
            "24-RG-BPBD-PROVSU_METRO.png"
        )
    return match.group("project").strip(), match.group("link").upper()


def resolve_tesseract(settings: Settings) -> str:
    configured = Path(settings.tesseract_command)
    if configured.exists():
        return str(configured)
    discovered = shutil.which(settings.tesseract_command) or shutil.which("tesseract")
    if discovered:
        return discovered
    raise RuntimeError(
        "Tesseract OCR tidak ditemukan. Instal Tesseract lalu periksa TESSERACT_COMMAND di .env."
    )


def run_ocr(image: Image.Image, settings: Settings) -> list[dict[str, str]]:
    prepared = ImageOps.grayscale(image)
    prepared = ImageEnhance.Contrast(prepared).enhance(2.2)
    prepared = prepared.resize((prepared.width * 2, prepared.height * 2))
    with tempfile.TemporaryDirectory(prefix="opd_ocr_") as directory:
        source = Path(directory) / "ocr.png"
        prepared.save(source)
        completed = subprocess.run(
            [resolve_tesseract(settings), str(source), "stdout", "--psm", "11", "tsv"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=45,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        if completed.returncode != 0:
            raise RuntimeError(f"Tesseract gagal: {completed.stderr.strip()[:300]}")
        return list(csv.DictReader(io.StringIO(completed.stdout), delimiter="\t"))


def extract_end_date_and_crop(path: Path, settings: Settings):
    from screenshot_reader import analyze
    with Image.open(path) as source_image:
        source = source_image.convert("RGB")
        info = analyze(run_ocr(source, settings), source.size)
        cropped = source.crop(info["box"])
        fd, name = tempfile.mkstemp(prefix="opd_cropped_", suffix=".png")
        os.close(fd)
        temp_path = Path(name)
        cropped.save(temp_path)
        return info["date"], temp_path, info["no_data"], info


def report_filename(report_date: date) -> str:
    return f"{report_date.day:02d} {MONTHS_ID[report_date.month]} {report_date.year}.xlsx"


def normalize_project(value: str) -> str:
    tokens = [token for token in re.split(r"[^A-Z0-9]+", value.upper().strip()) if token]
    return "".join(token for token in tokens if token != "RG")


def find_project_sheet(workbook: Any, project_name: str) -> Any:
    target = normalize_project(project_name)
    matches = [sheet for sheet in workbook.worksheets if normalize_project(sheet.title) == target]
    if len(matches) == 1:
        return matches[0]
    if not matches:
        raise ValueError(f"Sheet untuk project {project_name!r} tidak ditemukan.")
    raise ValueError(f"Project {project_name!r} cocok dengan lebih dari satu sheet.")


def project_sort_key(name: str) -> tuple:
    match = re.match(r"^\s*(\d+)", name)
    return (0, int(match.group(1)), name.strip().casefold()) if match else (1, 0, name.strip().casefold())


def load_project_choices(report: Path) -> list[str]:
    """Ambil nama project langsung dari seluruh nama sheet file laporan."""
    with zipfile.ZipFile(report) as archive:
        root = ET.fromstring(archive.read("xl/workbook.xml"))
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    return sorted([sheet.attrib["name"] for sheet in root.findall("s:sheets/s:sheet", ns)], key=project_sort_key)


def show_project_popup(projects: list[str], report: Path, suggested: tuple[str, str] = ("", "")) -> tuple[str, str] | None:
    """Grid project responsif tanpa dropdown atau scrollbar."""
    import math
    import tkinter as tk
    projects = sorted(projects, key=project_sort_key)
    if not projects:
        raise ValueError(f"Tidak ada nama sheet pada {report.name}.")
    root = tk.Tk()
    root.title("OPD — Konfirmasi laporan")
    root.configure(bg="#F1F5F9")
    root.attributes("-topmost", True)
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    width, height = min(1040, sw - 48), min(650, sh - 100)
    root.geometry(f"{width}x{height}+{max(0,(sw-width)//2)}+{max(0,(sh-height)//2)}")
    root.rowconfigure(1, weight=1)
    root.columnconfigure(0, weight=1)
    result = {}
    selected = tk.StringVar(value=suggested[0] if suggested[0] in projects else "")
    link = tk.StringVar(value=suggested[1] if suggested[1] in {"METRO", "BROADBAND"} else "")
    project_buttons = []
    network_buttons = []

    header = tk.Frame(root, bg="white", padx=20, pady=12)
    header.grid(row=0, column=0, sticky="ew")
    tk.Label(header, text="Konfirmasi laporan", bg="white", fg="#0F172A",
             font=("Segoe UI", 16, "bold")).pack(anchor="w")
    tk.Label(header, text=f"{report.name}   •   {len(projects)} project   •   Periksa hasil deteksi di bawah",
             bg="white", fg="#475569", font=("Segoe UI", 10)).pack(anchor="w", pady=(4,0))

    grid = tk.Frame(root, bg="#F1F5F9", padx=12, pady=8)
    grid.grid(row=1, column=0, sticky="nsew")
    footer = tk.Frame(root, bg="white", padx=20, pady=10)
    footer.grid(row=2, column=0, sticky="ew")
    footer.columnconfigure(0, weight=1)
    summary = tk.Label(footer, text="Pilih satu project dan jenis jaringan.",
                       bg="white", fg="#475569", font=("Segoe UI", 10), anchor="w")
    summary.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0,10))
    networks = tk.Frame(footer, bg="white")
    networks.grid(row=1, column=0, sticky="w")

    def refresh():
        for name, button in project_buttons:
            active = name == selected.get()
            button.configure(bg="#E0E7FF" if active else "white",
                             fg="#312E81" if active else "#1E293B",
                             text=("✓  " if active else "   ") + name.strip(),
                             relief="flat")
        for name, button in network_buttons:
            active = name == link.get()
            button.configure(bg="#4338CA" if active else "#F1F5F9",
                             fg="white" if active else "#334155")
        ready = bool(selected.get() and link.get())
        submit_button.configure(state="normal" if ready else "disabled",
                                bg="#4338CA" if ready else "#E2E8F0")
        summary.configure(text=(f"{selected.get().strip()}   /   {link.get() or 'Pilih jaringan'}"
                                if selected.get() else "Pilih satu project dan jenis jaringan."))

    def choose_project(name):
        selected.set(name)
        refresh()

    for name in projects:
        button = tk.Button(grid, text=name.strip(), anchor="w", justify="left",
                           bg="white", fg="#1E293B", activebackground="#EEF2FF",
                           activeforeground="#312E81", relief="flat", bd=0,
                           highlightthickness=1, highlightbackground="#E2E8F0",
                           highlightcolor="#4338CA",
                           padx=8, pady=3, font=("Segoe UI", 10),
                           cursor="hand2", command=lambda n=name: choose_project(n))
        project_buttons.append((name,button))

    def choose_link(name):
        link.set(name)
        refresh()

    for name, label in [("METRO","Metro"),("BROADBAND","Broadband")]:
        button = tk.Button(networks, text=label, font=("Segoe UI",11,"bold"),
                           bg="#F1F5F9", fg="#334155", relief="flat",
                           padx=14, pady=7, cursor="hand2",
                           command=lambda n=name: choose_link(n))
        button.pack(side="left", padx=(0,8))
        network_buttons.append((name,button))

    def submit():
        if selected.get() in projects and link.get() in {"METRO","BROADBAND"}:
            result.update(project=selected.get(), link=link.get())
            root.destroy()

    tk.Button(footer, text="Batal", command=root.destroy, bg="#E2E8F0",
              fg="#334155", relief="flat", padx=18, pady=9,
              font=("Segoe UI",11)).grid(row=1,column=1,padx=10)
    submit_button = tk.Button(footer, text="Konfirmasi & simpan", command=submit,
                              state="disabled", bg="#4338CA", fg="white",
                              disabledforeground="#475569", relief="flat",
                              padx=16,pady=8,font=("Segoe UI",11,"bold"))
    submit_button.grid(row=1,column=2)
    layout_state = [None]
    def layout(event=None):
        w, h = grid.winfo_width()-24, grid.winfo_height()-16
        if w < 100 or h < 100:
            return
        cols = max(3, min(6, math.ceil(len(projects) / max(1, h//44))))
        rows = math.ceil(len(projects)/cols)
        marker = (w,h,cols)
        if layout_state[0] == marker:
            return
        layout_state[0] = marker
        for c in range(6):
            grid.columnconfigure(c,weight=1 if c<cols else 0, uniform="cards" if c<cols else "")
        for r in range(len(projects)):
            grid.rowconfigure(r,weight=1 if r<rows else 0)
        for i, (_,button) in enumerate(project_buttons):
            button.grid(row=i//cols,column=i%cols,sticky="nsew",padx=4,pady=3)
            button.configure(wraplength=max(80,w//cols-32))
    grid.bind("<Configure>",layout)
    root.bind("<Escape>",lambda e:root.destroy())
    root.bind("<Return>",lambda e:submit())
    refresh()
    root.mainloop()
    return (result["project"],result["link"]) if result else None


def find_link_anchor(sheet: Any, link_type: str) -> str:
    targets = {"BROADBAND": "A2", "METRO": "A18"}
    if link_type not in targets:
        raise ValueError(f"Jenis jaringan tidak dikenal: {link_type}")
    return targets[link_type]


def anchor_position(image: Any) -> tuple[int, int] | None:
    marker = getattr(image.anchor, "_from", None)
    return None if marker is None else (marker.col + 1, marker.row + 1)


def protect_or_clear_slot(sheet: Any, anchor: str, allow_replace: bool) -> None:
    row, column = coordinate_to_tuple(anchor)
    # Slot versi sebelumnya juga dilindungi agar gambar lama tidak bertumpuk.
    positions = {(column, row), (column, row - 1)}
    existing = [image for image in sheet._images if anchor_position(image) in positions]
    if existing and not allow_replace:
        raise FileExistsError(
            f"Slot {sheet.title}!{anchor} sudah memiliki gambar. ALLOW_REPLACE_EXISTING=false."
        )
    if existing:
        sheet._images = [image for image in sheet._images if image not in existing]


def calculate_size(path: Path, settings: Settings) -> tuple[int, int]:
    with Image.open(path) as image:
        width, height = image.size
    scale = min(settings.max_image_width_px / width, settings.max_image_height_px / height, 1.0)
    return max(1, round(width * scale)), max(1, round(height * scale))


def create_backup(report: Path, settings: Settings) -> Path:
    folder = settings.backup_folder / datetime.now().strftime("%Y-%m-%d")
    folder.mkdir(parents=True, exist_ok=True)
    destination = folder / f"{report.stem}_{datetime.now():%H%M%S_%f}.xlsx"
    shutil.copy2(report, destination)
    return destination


def insert_into_drive_report(
    path: Path,
    settings: Settings,
    selection: tuple[str, str] | None = None,
) -> dict[str, Any] | None:
    with timed("OCR tanggal dan pemotongan"):
        report_date, cropped_path, no_data, detected = extract_end_date_and_crop(path, settings)
    image_path = cropped_path or path
    report = settings.drive_report_folder / report_filename(report_date)
    try:
        if not report.exists():
            raise FileNotFoundError(f"File laporan tidak ditemukan: {report}")
        if selection is None:
            from screenshot_reader import suggest_selection
            projects = load_project_choices(report)
            suggested = suggest_selection(detected, projects)
            logging.info("Saran OCR: %s | %s. Menunggu konfirmasi popup.",
                         suggested[0] or "belum terbaca", suggested[1] or "belum terbaca")
            selection = show_project_popup(projects, report, suggested)
            if selection is None:
                logging.info("Pemilihan dibatalkan: %s", path.name)
                return None
        project, link_type = selection
        baseline = file_version(report)
        with timed("Membuka Excel"):
            workbook = load_workbook(report)
        sheet = find_project_sheet(workbook, project)
        anchor = find_link_anchor(sheet, link_type)
        protect_or_clear_slot(sheet, anchor, settings.allow_replace_existing)

        width, height = calculate_size(image_path, settings)
        excel_image = ExcelImage(str(image_path))
        excel_image.width, excel_image.height = width, height
        sheet.add_image(excel_image, anchor)

        with timed("Backup"):
            backup = create_backup(report, settings)
        fd, temp_name = tempfile.mkstemp(prefix="opd_saving_", suffix=".xlsx")
        os.close(fd)
        temp_path = Path(temp_name)
        try:
            with timed("Menyimpan Excel di disk lokal"):
                workbook.save(temp_path)
            with timed("Verifikasi arsip Excel"):
                with zipfile.ZipFile(temp_path) as archive:
                    bad = archive.testzip()
                    if bad:
                        raise ValueError(f"Arsip Excel rusak: {bad}")
                    ET.fromstring(archive.read("xl/workbook.xml"))
            with timed("Menyalin hasil ke folder Drive"):
                if file_version(report) != baseline:
                    raise RuntimeError("File tujuan berubah selama proses. Ulangi screenshot agar perubahan tidak tertimpa.")
                transfer_fd, transfer_name = tempfile.mkstemp(
                    prefix="opd_transfer_", suffix=".xlsx", dir=report.parent)
                os.close(transfer_fd)
                transfer = Path(transfer_name)
                try:
                    shutil.copyfile(temp_path, transfer)
                    if file_version(report) != baseline:
                        raise RuntimeError("File tujuan berubah saat transfer. Silakan ulangi.")
                    os.replace(transfer, report)
                finally:
                    transfer.unlink(missing_ok=True)
            logging.info("Excel tersimpan. Status sinkronisasi online diperiksa melalui aplikasi Google Drive.")
        except PermissionError as exc:
            raise PermissionError(
                f"File {report.name} sedang dibuka di Excel. Tutup file lalu coba lagi."
            ) from exc
        finally:
            temp_path.unlink(missing_ok=True)

        return {
            "status": "no_data" if no_data else "normal",
            "project": project,
            "link_type": link_type,
            "report_date": report_date.isoformat(),
            "workbook": report.name,
            "workbook_path": str(report),
            "sheet": sheet.title,
            "cell": anchor,
            "backup": str(backup),
            "image": path.name,
            "image_path": str(path.resolve()),
        }
    finally:
        if cropped_path:
            cropped_path.unlink(missing_ok=True)


def notification_text(result: dict[str, Any]) -> str:
    return (
        "Screenshot berhasil dimasukkan ke laporan OPD.\n"
        f"Project: {result['project']}\nLink: {result['link_type']}\n"
        f"Tanggal laporan: {result['report_date']}\nFile: {result['workbook']}\n"
        f"Sheet: {result['sheet']}\nPosisi: {result['cell']}\nWaktu: {result['timestamp']}"
    )


def alert_popup(result: dict[str, Any]) -> None:
    import tkinter as tk
    from PIL import ImageTk
    from clipboard_helper import copy_content
    root = tk.Tk()
    root.title("OPD • No Data — Siap disalin ke WhatsApp")
    root.configure(bg="#F1F5F9")
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    w, h = min(1100, sw-40), min(760, sh-90)
    root.geometry(f"{w}x{h}+{max(0,(sw-w)//2)}+{max(0,(sh-h)//2)}")
    root.columnconfigure(0, weight=3)
    root.columnconfigure(1, weight=2)
    root.rowconfigure(1, weight=1)
    tk.Label(root, text="No Data terdeteksi • Laporan siap dibagikan",
             bg="#0F2942", fg="white", font=("Segoe UI", 17, "bold"),
             padx=20, pady=16, anchor="w").grid(row=0,column=0,columnspan=2,sticky="ew")
    preview = tk.Label(root, bg="#DDE5EE", text="Screenshot asli tidak tersedia.")
    preview.grid(row=1,column=0,sticky="nsew",padx=(16,8),pady=16)
    original = None
    try:
        with Image.open(result["image_path"]) as source:
            original = source.copy()
    except (OSError, KeyError) as exc:
        logging.warning("Tidak dapat membuka screenshot asli: %s", exc)
    def render(event=None):
        if original is None:
            return
        image = original.copy()
        image.thumbnail((max(1,preview.winfo_width()-16),
                         max(1,preview.winfo_height()-16)), Image.Resampling.LANCZOS)
        preview.photo = ImageTk.PhotoImage(image, master=root)
        preview.configure(image=preview.photo,text="")
    preview.bind("<Configure>",render)
    right = tk.Frame(root,bg="white",padx=16,pady=16)
    right.grid(row=1,column=1,sticky="nsew",padx=(8,16),pady=16)
    right.rowconfigure(1,weight=1)
    right.columnconfigure(0,weight=1)
    tk.Label(right,text="Pesan (bisa diedit sebelum disalin)",font=("Segoe UI",11,"bold"),
             bg="white",fg="#16324F").grid(row=0,column=0,sticky="w",pady=(0,10))
    message = tk.Text(right,wrap="word",font=("Segoe UI",11),relief="flat",
                      bg="#F8FAFC",fg="#16324F",padx=12,pady=12,width=32,height=8)
    message.grid(row=1,column=0,sticky="nsew")
    message.insert("1.0",(
        f"PERINGATAN NO DATA OPD\n\nProject: {result['project'].strip()}\n"
        f"Jaringan: {result['link_type']}\nTanggal laporan: {result['report_date']}\n\n"
        "Ruijie menampilkan No Data pada periode tersebut.\nMohon dilakukan pemeriksaan jaringan."
    ))
    status = tk.StringVar(value="Belum dikirim. Anda memilih penerima dan mengirim sendiri di WhatsApp.")
    def copy(mode):
        try:
            root.update_idletasks()
            copy_content(root.winfo_id(),
                         text=message.get("1.0","end-1c") if mode != "image" else None,
                         image_path=result["image_path"] if mode != "text" else None)
            if mode == "both":
                status.set("Pesan dan gambar disalin. Tempel di WhatsApp. Jika caption kosong, klik Salin Pesan lalu tempel pada caption.")
            else:
                status.set("Pesan disalin." if mode == "text" else "Screenshot asli utuh disalin.")
        except Exception as exc:
            status.set(f"Gagal menyalin: {exc}")
    tk.Button(right,text="Salin Pesan + Screenshot",command=lambda:copy("both"),
              state="normal" if original else "disabled",bg="#126B69",fg="white",
              font=("Segoe UI",11,"bold"),relief="flat",pady=12).grid(row=2,column=0,sticky="ew",pady=(16,6))
    extra = tk.Frame(right,bg="white")
    extra.grid(row=3,column=0,sticky="ew")
    tk.Button(extra,text="Salin Pesan",command=lambda:copy("text"),pady=8).pack(side="left",fill="x",expand=True,padx=(0,4))
    tk.Button(extra,text="Salin Screenshot",command=lambda:copy("image"),
              state="normal" if original else "disabled",pady=8).pack(side="left",fill="x",expand=True)
    tk.Label(root,textvariable=status,bg="#F1F5F9",fg="#334155",wraplength=w-60,
             font=("Segoe UI",10),padx=20,pady=8,justify="left").grid(row=2,column=0,columnspan=2,sticky="ew")
    tk.Button(root,text="Tutup",command=root.destroy,padx=24,pady=8).grid(
        row=3,column=0,columnspan=2,pady=(0,12))
    root.bind("<Escape>",lambda e:root.destroy())
    root.lift()
    root.mainloop()


def send_notification(settings: Settings, result: dict[str, Any]) -> None:
    # Legacy .env mode/credentials intentionally have no effect: no network sends.
    if result.get("status") != "no_data":
        logging.info("Screenshot tersimpan. Tidak ada peringatan No Data.")
        return
    logging.info("NO DATA: %s. Membuka popup salin manual.", result["project"])
    alert_popup(result)


def process_file(path: Path, settings: Settings, state: StateStore) -> None:
    path = path.resolve()
    if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        return
    if not path.exists() or not wait_until_stable(path):
        logging.warning("File belum stabil: %s", path)
        return
    key = fingerprint(path)
    if state.was_processed(key):
        logging.info("Sudah pernah diproses: %s", path.name)
        return
    try:
        result = insert_into_drive_report(path, settings)
        if result is None:
            return
        result["timestamp"] = local_now(settings).isoformat(timespec="seconds")
        state.mark_processed(key, result)
        logging.info(
            "Berhasil: %s -> %s | %s!%s",
            path.name, result["workbook"], result["sheet"], result["cell"],
        )
        try:
            send_notification(settings, result)
        except Exception:
            logging.exception("Gambar masuk ke Drive, tetapi notifikasi gagal.")
    except ValueError as exc:
        logging.warning("Screenshot perlu diperiksa: %s", exc)
    except FileExistsError as exc:
        logging.warning(
            "NOTIFIKASI DUPLIKAT / SLOT SUDAH TERISI\n"
            "Screenshot tidak dimasukkan dan file Excel tidak diubah.\n"
            "Screenshot: %s\nAlasan: %s",
            path.name,
            exc,
        )
    except Exception:
        logging.exception("Gagal memproses screenshot: %s", path.name)


def watch(settings: Settings, state: StateStore) -> None:
    if not settings.screenshot_folder.exists():
        raise FileNotFoundError(f"Folder screenshot tidak ditemukan: {settings.screenshot_folder}")
    if not settings.drive_report_folder.exists():
        raise FileNotFoundError(f"Folder Drive tidak ditemukan: {settings.drive_report_folder}")
    resolve_tesseract(settings)

    logging.info("Memantau screenshot: %s", settings.screenshot_folder)
    logging.info("Folder laporan Drive: %s", settings.drive_report_folder)
    logging.info("Mode notifikasi: popup salin manual (tanpa pengiriman otomatis)")
    logging.info("Tekan Ctrl+C untuk berhenti.")
    pending = queue.Queue()
    stop = threading.Event()

    def scan():
        result = {}
        for path in settings.screenshot_folder.iterdir():
            if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS:
                try:
                    result[path.resolve()] = file_version(path)
                except OSError:
                    continue
        return result

    initial = scan()
    def collect():
        seen = initial
        while not stop.wait(0.4):
            try:
                current = scan()
                for path, signature in current.items():
                    if seen.get(path) != signature:
                        pending.put(path)
                        logging.info("Antrean screenshot: %s (menunggu %s)", path.name, pending.qsize())
                seen = current
            except OSError as exc:
                logging.warning("Folder belum dapat dibaca: %s", exc)

    worker = threading.Thread(target=collect, daemon=True)
    worker.start()
    try:
        while True:
            try:
                path = pending.get(timeout=0.4)
            except queue.Empty:
                continue
            try:
                process_file(path, settings, state)
            finally:
                pending.task_done()
    except KeyboardInterrupt:
        logging.info("Program dihentikan. Antrean belum selesai: %s. Gunakan --once untuk mengulang file lama.", pending.qsize())
    finally:
        stop.set()
        worker.join(timeout=2)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Masukkan screenshot OPD ke Excel di Google Drive.")
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--watch", action="store_true")
    group.add_argument("--once", type=Path)
    return parser.parse_args()


def main() -> int:
    setup_logging()
    logging.info("OPD v14 — deteksi project dan konfirmasi otomatis")
    try:
        settings = Settings.from_env()
        state = StateStore(PROJECT_DIR / "data" / "state.json")
        args = parse_args()
        process_file(args.once, settings, state) if args.once else watch(settings, state)
        return 0
    except Exception:
        logging.exception("Program berhenti karena kesalahan konfigurasi.")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
