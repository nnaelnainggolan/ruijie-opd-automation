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
import urllib.error
import urllib.request
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
        )
        if completed.returncode != 0:
            raise RuntimeError(f"Tesseract gagal: {completed.stderr.strip()[:300]}")
        return list(csv.DictReader(io.StringIO(completed.stdout), delimiter="\t"))


def extract_end_date_and_crop(path: Path, settings: Settings) -> tuple[date, Path | None]:
    with Image.open(path) as source_image:
        source = source_image.convert("RGB")
        words = run_ocr(source, settings)
        text = " ".join(row.get("text", "") for row in words if row.get("text", "").strip())
        match = DATE_RANGE_PATTERN.search(re.sub(r"\s+", "", text).upper())
        if not match:
            raise ValueError(f"Tanggal tidak terbaca dari screenshot. Hasil OCR: {text[:180]}")
        end_date = date(int(match["y2"]), int(match["m2"]), int(match["d2"]))
        if not settings.crop_speed_summary:
            return end_date, None

        y_values = []
        for row in words:
            token = row.get("text", "").strip().lower()
            if token in {"speed", "summary"} or re.search(r"20\d{2}[/\-.]\d{1,2}", token):
                try:
                    y_values.append(int(row["top"]) // 2)
                except (KeyError, TypeError, ValueError):
                    pass
        crop_top = max(0, min(y_values) - 12) if y_values else 0
        cropped = source.crop((0, crop_top, source.width, source.height))
        fd, temp_name = tempfile.mkstemp(prefix="opd_cropped_", suffix=".png")
        os.close(fd)
        temp_path = Path(temp_name)
        cropped.save(temp_path)
        return end_date, temp_path


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


def load_project_choices(report: Path) -> list[str]:
    """Ambil nama project langsung dari seluruh nama sheet file laporan."""
    workbook = load_workbook(report, read_only=True, data_only=True)
    try:
        return [name.strip() for name in workbook.sheetnames if name.strip()]
    finally:
        workbook.close()


def show_project_popup(projects: list[str], report: Path) -> tuple[str, str] | None:
    """Tampilkan pilihan project dan jenis jaringan pada desktop Windows."""
    try:
        import tkinter as tk
        from tkinter import messagebox, ttk
    except ImportError as exc:
        raise RuntimeError("Tkinter tidak tersedia pada instalasi Python ini.") from exc

    if not projects:
        raise ValueError(f"Tidak ada nama sheet pada {report.name}.")

    result: dict[str, str] = {}
    root = tk.Tk()
    root.title("Pilih Tujuan Screenshot OPD")
    root.resizable(False, False)
    root.attributes("-topmost", True)

    frame = ttk.Frame(root, padding=20)
    frame.grid(row=0, column=0, sticky="nsew")
    ttk.Label(frame, text="Screenshot baru terdeteksi", font=("Segoe UI", 12, "bold")).grid(
        row=0, column=0, columnspan=2, sticky="w", pady=(0, 5)
    )
    ttk.Label(frame, text=f"File laporan: {report.name}").grid(
        row=1, column=0, columnspan=2, sticky="w", pady=(0, 16)
    )
    ttk.Label(frame, text="Nama project").grid(row=2, column=0, sticky="w", pady=5)
    project_var = tk.StringVar(value=projects[0])
    project_box = ttk.Combobox(
        frame, textvariable=project_var, values=projects, state="readonly", width=42
    )
    project_box.grid(row=2, column=1, sticky="ew", pady=5)

    ttk.Label(frame, text="Jenis jaringan").grid(row=3, column=0, sticky="nw", pady=8)
    link_var = tk.StringVar(value="METRO")
    link_frame = ttk.Frame(frame)
    link_frame.grid(row=3, column=1, sticky="w", pady=5)
    ttk.Radiobutton(link_frame, text="Metro", variable=link_var, value="METRO").pack(
        side="left", padx=(0, 18)
    )
    ttk.Radiobutton(
        link_frame, text="Broadband", variable=link_var, value="BROADBAND"
    ).pack(side="left")

    def submit() -> None:
        if not project_var.get().strip():
            messagebox.showwarning("Project belum dipilih", "Silakan pilih nama project.")
            return
        result["project"] = project_var.get().strip()
        result["link"] = link_var.get()
        root.destroy()

    def cancel() -> None:
        root.destroy()

    buttons = ttk.Frame(frame)
    buttons.grid(row=4, column=0, columnspan=2, sticky="e", pady=(18, 0))
    ttk.Button(buttons, text="Batal", command=cancel).pack(side="left", padx=(0, 8))
    ttk.Button(buttons, text="Proses", command=submit).pack(side="left")
    root.protocol("WM_DELETE_WINDOW", cancel)
    root.bind("<Escape>", lambda _event: cancel())
    root.bind("<Return>", lambda _event: submit())
    root.update_idletasks()
    x = max(0, (root.winfo_screenwidth() - root.winfo_reqwidth()) // 2)
    y = max(0, (root.winfo_screenheight() - root.winfo_reqheight()) // 2)
    root.geometry(f"+{x}+{y}")
    project_box.focus_set()
    root.mainloop()
    return (result["project"], result["link"]) if result else None


def find_link_anchor(sheet: Any, link_type: str) -> str:
    expected = f"LINK {link_type}"
    for row in sheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.strip().upper().startswith(expected):
                return cell.coordinate
    raise ValueError(f"Judul {expected} tidak ditemukan di sheet {sheet.title!r}.")


def anchor_position(image: Any) -> tuple[int, int] | None:
    marker = getattr(image.anchor, "_from", None)
    return None if marker is None else (marker.col + 1, marker.row + 1)


def protect_or_clear_slot(sheet: Any, anchor: str, allow_replace: bool) -> None:
    row, column = coordinate_to_tuple(anchor)
    existing = [image for image in sheet._images if anchor_position(image) == (column, row)]
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
    destination = folder / f"{report.stem}_{datetime.now():%H%M%S}.xlsx"
    shutil.copy2(report, destination)
    return destination


def insert_into_drive_report(
    path: Path,
    settings: Settings,
    selection: tuple[str, str] | None = None,
) -> dict[str, Any] | None:
    report_date, cropped_path = extract_end_date_and_crop(path, settings)
    image_path = cropped_path or path
    report = settings.drive_report_folder / report_filename(report_date)
    try:
        if not report.exists():
            raise FileNotFoundError(f"File laporan tidak ditemukan: {report}")
        if selection is None:
            try:
                project, link_type = parse_screenshot_filename(path)
            except ValueError:
                selection = show_project_popup(load_project_choices(report), report)
                if selection is None:
                    logging.info("Pemilihan dibatalkan: %s", path.name)
                    return None
                project, link_type = selection
        else:
            project, link_type = selection
        workbook = load_workbook(report)
        sheet = find_project_sheet(workbook, project)
        anchor = find_link_anchor(sheet, link_type)
        protect_or_clear_slot(sheet, anchor, settings.allow_replace_existing)

        width, height = calculate_size(image_path, settings)
        excel_image = ExcelImage(str(image_path))
        excel_image.width, excel_image.height = width, height
        sheet.add_image(excel_image, anchor)

        backup = create_backup(report, settings)
        fd, temp_name = tempfile.mkstemp(prefix="opd_saving_", suffix=".xlsx", dir=report.parent)
        os.close(fd)
        temp_path = Path(temp_name)
        try:
            workbook.save(temp_path)
            verification = load_workbook(temp_path)
            verification.close()
            os.replace(temp_path, report)
        except PermissionError as exc:
            raise PermissionError(
                f"File {report.name} sedang dibuka di Excel. Tutup file lalu coba lagi."
            ) from exc
        finally:
            temp_path.unlink(missing_ok=True)

        return {
            "project": project,
            "link_type": link_type,
            "report_date": report_date.isoformat(),
            "workbook": report.name,
            "workbook_path": str(report),
            "sheet": sheet.title,
            "cell": anchor,
            "backup": str(backup),
            "image": path.name,
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


def post_json(url: str, payload: dict[str, Any], headers: dict[str, str]) -> None:
    request = urllib.request.Request(
        url, data=json.dumps(payload, ensure_ascii=False).encode(), method="POST"
    )
    request.add_header("Content-Type", "application/json")
    for name, value in headers.items():
        request.add_header(name, value)
    try:
        with urllib.request.urlopen(request, timeout=20) as response:
            if not 200 <= response.status < 300:
                raise RuntimeError(f"HTTP {response.status}")
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode(errors="replace")[:500]
        raise RuntimeError(f"HTTP {exc.code}: {detail}") from exc


def send_notification(settings: Settings, result: dict[str, Any]) -> None:
    mode = settings.notification_mode
    payload = {
        "event": "opd_screenshot_inserted", "status": "success",
        "message": "Screenshot berhasil dimasukkan ke laporan OPD.", **result,
    }
    if mode == "console":
        logging.info("NOTIFIKASI\n%s", notification_text(result))
    elif mode == "n8n":
        if not settings.n8n_webhook_url:
            raise RuntimeError("N8N_WEBHOOK_URL belum diisi.")
        headers = ({"Authorization": f"Bearer {settings.n8n_webhook_token}"}
                   if settings.n8n_webhook_token else {})
        post_json(settings.n8n_webhook_url, payload, headers)
        logging.info("Notifikasi berhasil dikirim ke n8n.")
    elif mode == "whatsapp_cloud":
        required = {
            "WHATSAPP_ACCESS_TOKEN": settings.whatsapp_access_token,
            "WHATSAPP_PHONE_NUMBER_ID": settings.whatsapp_phone_number_id,
            "WHATSAPP_TO": settings.whatsapp_to,
        }
        missing = [key for key, value in required.items() if not value]
        if missing:
            raise RuntimeError(f"Konfigurasi WhatsApp belum lengkap: {', '.join(missing)}")
        url = (f"https://graph.facebook.com/{settings.whatsapp_api_version}/"
               f"{settings.whatsapp_phone_number_id}/messages")
        body = {
            "messaging_product": "whatsapp", "recipient_type": "individual",
            "to": settings.whatsapp_to, "type": "text",
            "text": {"preview_url": False, "body": notification_text(result)},
        }
        post_json(url, body, {"Authorization": f"Bearer {settings.whatsapp_access_token}"})
        logging.info("Notifikasi berhasil dikirim melalui WhatsApp Cloud API.")
    else:
        raise RuntimeError("NOTIFICATION_MODE harus console, n8n, atau whatsapp_cloud.")


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
    logging.info("Mode notifikasi: %s", settings.notification_mode)
    logging.info("Tekan Ctrl+C untuk berhenti.")
    seen = {
        path.resolve() for path in settings.screenshot_folder.iterdir()
        if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS
    }
    try:
        while True:
            current = {
                path.resolve() for path in settings.screenshot_folder.iterdir()
                if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS
            }
            for path in sorted(current - seen, key=lambda item: item.stat().st_mtime_ns):
                process_file(path, settings, state)
            seen = current
            time.sleep(1)
    except KeyboardInterrupt:
        logging.info("Program dihentikan.")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Masukkan screenshot OPD ke Excel di Google Drive.")
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--watch", action="store_true")
    group.add_argument("--once", type=Path)
    return parser.parse_args()


def main() -> int:
    setup_logging()
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
