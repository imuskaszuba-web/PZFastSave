import os
import shutil
import subprocess
import sys
import tarfile
import time
import zipfile
from dataclasses import dataclass
from pathlib import Path
from zipfile import ZipInfo

try:
    from PyQt6.QtCore import Qt, QThread, QTimer, pyqtSignal
    from PyQt6.QtWidgets import (
        QApplication,
        QCheckBox,
        QComboBox,
        QFileDialog,
        QGridLayout,
        QGroupBox,
        QHBoxLayout,
        QHeaderView,
        QLabel,
        QLineEdit,
        QMainWindow,
        QMessageBox,
        QPushButton,
        QProgressBar,
        QTableWidget,
        QTableWidgetItem,
        QTextEdit,
        QVBoxLayout,
        QWidget,
    )
except ModuleNotFoundError as exc:
    if exc.name == "PyQt6":
        print("Missing dependency: PyQt6")
        print("Install it with: python -m pip install pyqt6")
        raise SystemExit(1)
    raise


DEFAULT_SAVE_ROOT = Path.home() / "Zomboid" / "Saves"
DEFAULT_BACKUP_ROOT = Path.home() / "Zomboid" / "SaveBackups"

ARCHIVE_KINDS = {
    ".zip": "zip",
    ".7z": "7z",
    ".wim": "wim",
    ".tar": "tar",
}


@dataclass(frozen=True)
class SaveEntry:
    mode: str
    name: str
    path: Path
    updated: float


@dataclass(frozen=True)
class BackupEntry:
    mode: str
    name: str
    path: Path
    updated: float
    kind: str


def format_time(timestamp: float) -> str:
    return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(timestamp))


def safe_name(value: str) -> str:
    invalid = '<>:"/\\|?*'
    cleaned = "".join("_" if ch in invalid else ch for ch in value)
    return cleaned.strip().rstrip(".") or "save"


def subprocess_no_window_kwargs() -> dict:
    if os.name == "nt":
        return {"creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0)}
    return {}


def find_local_or_path(executable_names: tuple[str, ...]) -> str | None:
    app_dir = Path(__file__).resolve().parent

    for name in executable_names:
        local = app_dir / name
        if local.exists() and local.is_file():
            return str(local)

    for name in executable_names:
        found = shutil.which(name)
        if found:
            return found

    return None


def find_7zip() -> str | None:
    found = find_local_or_path(("7z.exe", "7zz.exe", "7za.exe", "7z", "7zz", "7za"))
    if found:
        return found

    if os.name == "nt":
        for value in (
            r"C:\Program Files\7-Zip\7z.exe",
            r"C:\Program Files (x86)\7-Zip\7z.exe",
        ):
            path = Path(value)
            if path.exists() and path.is_file():
                return str(path)

    return None


def require_7zip() -> str:
    executable = find_7zip()
    if not executable:
        raise RuntimeError("7-Zip not found. Install 7-Zip or put 7z.exe in PATH.")
    return executable


def find_wimlib() -> str | None:
    return find_local_or_path(("wimlib-imagex.exe", "wimlib-imagex"))


def require_wimlib() -> str:
    executable = find_wimlib()
    if not executable:
        raise RuntimeError(
            "wimlib-imagex.exe not found. Put wimlib-imagex.exe next to this script or add it to PATH."
        )
    return executable


def backup_kind_from_path(path: Path) -> str | None:
    if path.is_dir():
        return "folder"

    if not path.is_file():
        return None

    lower_name = path.name.lower()

    if lower_name.endswith(".tar.gz"):
        return "tar.gz"

    return ARCHIVE_KINDS.get(path.suffix.lower())


def parse_backup_name(path: Path) -> tuple[str, str] | None:
    if path.is_dir():
        stem = path.name
    elif path.is_file():
        lower_name = path.name.lower()

        if lower_name.endswith(".tar.gz"):
            stem = path.name[:-7]
        elif path.suffix.lower() in ARCHIVE_KINDS:
            stem = path.stem
        else:
            return None
    else:
        return None

    parts = stem.split("__", 2)
    if len(parts) != 3:
        return None

    return parts[1], parts[2]


def scan_saves(save_root: Path) -> list[SaveEntry]:
    if not save_root.exists():
        return []

    saves: list[SaveEntry] = []

    for mode_dir in save_root.iterdir():
        if not mode_dir.is_dir():
            continue

        for save_dir in mode_dir.iterdir():
            if not save_dir.is_dir():
                continue

            try:
                updated = save_dir.stat().st_mtime
            except OSError:
                updated = 0

            saves.append(SaveEntry(mode_dir.name, save_dir.name, save_dir, updated))

    saves.sort(key=lambda item: item.updated, reverse=True)
    return saves


def scan_backups(backup_root: Path) -> list[BackupEntry]:
    if not backup_root.exists():
        return []

    backups: list[BackupEntry] = []

    for path in backup_root.iterdir():
        kind = backup_kind_from_path(path)
        if not kind:
            continue

        parsed = parse_backup_name(path)
        if not parsed:
            continue

        try:
            updated = path.stat().st_mtime
        except OSError:
            updated = 0

        mode, name = parsed
        backups.append(BackupEntry(mode, name, path, updated, kind))

    backups.sort(key=lambda item: item.updated, reverse=True)
    return backups


def unique_destination(root: Path, mode: str, name: str, suffix: str = "") -> Path:
    stamp = time.strftime("%Y%m%d-%H%M%S")
    stem = f"{stamp}__{safe_name(mode)}__{safe_name(name)}"
    candidate = root / f"{stem}{suffix}"

    counter = 2
    while candidate.exists():
        candidate = root / f"{stem}_{counter}{suffix}"
        counter += 1

    return candidate


def unique_sidecar_path(path: Path, marker: str) -> Path:
    stamp = time.strftime("%Y%m%d-%H%M%S")
    candidate = path.with_name(f"{path.name}__{marker}__{stamp}")

    counter = 2
    while candidate.exists():
        candidate = path.with_name(f"{path.name}__{marker}__{stamp}_{counter}")
        counter += 1

    return candidate


def run_robocopy(source: Path, destination: Path) -> None:
    destination.mkdir(parents=True, exist_ok=True)

    command = [
        "robocopy",
        str(source),
        str(destination),
        "/E",
        "/MT:64",
        "/R:0",
        "/W:0",
        "/XJ",
        "/COPY:DAT",
        "/DCOPY:DAT",
        "/NFL",
        "/NDL",
        "/NJH",
        "/NJS",
        "/NP",
    ]

    result = subprocess.run(
        command,
        capture_output=True,
        text=True,
        **subprocess_no_window_kwargs(),
    )

    if result.returncode >= 8:
        output = (result.stdout + result.stderr).strip()
        raise RuntimeError(output or f"robocopy failed: {result.returncode}")


def copy_tree_fast(source: Path, destination: Path) -> None:
    if os.name == "nt" and shutil.which("robocopy"):
        run_robocopy(source, destination)
        return

    shutil.copytree(source, destination, copy_function=shutil.copy2)


def zip_store_fast(source: Path, destination_zip: Path) -> None:
    seven_zip = find_7zip()
    temp_zip = destination_zip.with_name(destination_zip.name + ".tmp")

    if temp_zip.exists():
        temp_zip.unlink()

    if seven_zip:
        command = [
            seven_zip,
            "a",
            "-tzip",
            "-mx=0",
            "-mm=Copy",
            "-mmt=on",
            "-y",
            "-bb0",
            "-bso0",
            "-bsp0",
            str(temp_zip),
            source.name,
        ]

        result = subprocess.run(
            command,
            cwd=source.parent,
            capture_output=True,
            text=True,
            **subprocess_no_window_kwargs(),
        )

        if result.returncode != 0:
            if temp_zip.exists():
                temp_zip.unlink()

            output = (result.stdout + result.stderr).strip()
            raise RuntimeError(output or f"7-Zip zip failed: {result.returncode}")

        temp_zip.replace(destination_zip)
        return

    try:
        with zipfile.ZipFile(
            temp_zip,
            "w",
            compression=zipfile.ZIP_STORED,
            allowZip64=True,
        ) as archive:
            for file_path in source.rglob("*"):
                if not file_path.is_file():
                    continue

                archive.write(file_path, file_path.relative_to(source.parent))

        temp_zip.replace(destination_zip)

    except Exception:
        if temp_zip.exists():
            temp_zip.unlink()
        raise


def seven_zip_fast_full_backup(source: Path, destination_7z: Path) -> None:
    seven_zip = require_7zip()
    temp_7z = destination_7z.with_name(destination_7z.name + ".tmp")

    if temp_7z.exists():
        temp_7z.unlink()

    threads = os.cpu_count() or 16

    command = [
        seven_zip,
        "a",
        "-t7z",
        "-mx=0",
        "-m0=Copy",
        "-ms=off",
        f"-mmt={threads}",
        "-y",
        "-bb0",
        "-bso0",
        "-bsp0",
        str(temp_7z),
        source.name,
    ]

    result = subprocess.run(
        command,
        cwd=source.parent,
        capture_output=True,
        text=True,
        **subprocess_no_window_kwargs(),
    )

    if result.returncode != 0:
        if temp_7z.exists():
            temp_7z.unlink()

        output = (result.stdout + result.stderr).strip()
        raise RuntimeError(output or f"7-Zip failed: {result.returncode}")

    temp_7z.replace(destination_7z)


def wim_full_backup_fastest(source: Path, destination_wim: Path) -> None:
    wimlib = require_wimlib()
    temp_wim = destination_wim.with_name(destination_wim.name + ".tmp")

    if temp_wim.exists():
        temp_wim.unlink()

    command = [
        wimlib,
        "capture",
        str(source),
        str(temp_wim),
        "ProjectZomboidSave",
        "--compress=none",
        "--threads=0",
        "--no-acls",
        "--norpfix",
    ]

    result = subprocess.run(
        command,
        capture_output=True,
        text=True,
        **subprocess_no_window_kwargs(),
    )

    if result.returncode != 0:
        if temp_wim.exists():
            temp_wim.unlink()

        output = (result.stdout + result.stderr).strip()
        raise RuntimeError(output or f"wimlib capture failed: {result.returncode}")

    temp_wim.replace(destination_wim)


def extract_archive_to_temp_root(source_archive: Path, destination: Path) -> Path:
    seven_zip = require_7zip()
    temp_extract_root = destination.with_name(destination.name + "__extracting")

    if temp_extract_root.exists():
        shutil.rmtree(temp_extract_root)

    temp_extract_root.mkdir(parents=True, exist_ok=True)

    command = [
        seven_zip,
        "x",
        "-y",
        "-aoa",
        "-mmt=on",
        "-bb0",
        "-bso0",
        "-bsp0",
        f"-o{temp_extract_root}",
        str(source_archive),
    ]

    result = subprocess.run(
        command,
        capture_output=True,
        text=True,
        **subprocess_no_window_kwargs(),
    )

    if result.returncode != 0:
        shutil.rmtree(temp_extract_root, ignore_errors=True)
        output = (result.stdout + result.stderr).strip()
        raise RuntimeError(output or f"7-Zip extract failed: {result.returncode}")

    children = [item for item in temp_extract_root.iterdir()]

    if len(children) == 1 and children[0].is_dir():
        return children[0]

    return temp_extract_root


def replace_destination_from_extracted_root(extracted_root: Path, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)

    if destination.exists():
        raise RuntimeError(f"Restore target already exists: {destination}")

    if extracted_root.name == destination.name:
        extracted_root.replace(destination)
        parent = extracted_root.parent
        if parent.exists() and parent.name.endswith("__extracting"):
            shutil.rmtree(parent, ignore_errors=True)
        return

    destination.mkdir(parents=True, exist_ok=True)

    for item in extracted_root.iterdir():
        item.replace(destination / item.name)

    parent = extracted_root.parent
    if parent.exists() and parent.name.endswith("__extracting"):
        shutil.rmtree(parent, ignore_errors=True)


def extract_7z_save_contents(source_archive: Path, destination: Path) -> None:
    extracted_root = extract_archive_to_temp_root(source_archive, destination)
    replace_destination_from_extracted_root(extracted_root, destination)


def extract_zip_save_contents(source_zip: Path, destination: Path) -> None:
    if find_7zip():
        extract_7z_save_contents(source_zip, destination)
        return

    destination.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(source_zip, "r") as archive:
        members = archive.infolist()

        top_level = None
        for member in members:
            parts = validate_zip_member(member)
            if parts:
                top_level = parts[0]
                break

        for member in members:
            parts = validate_zip_member(member)
            if not parts:
                continue

            if top_level and len(parts) > 1 and parts[0] == top_level:
                parts = parts[1:]

            if not parts:
                continue

            target = destination.joinpath(*parts)

            if member.is_dir():
                target.mkdir(parents=True, exist_ok=True)
                continue

            target.parent.mkdir(parents=True, exist_ok=True)

            with archive.open(member, "r") as src, target.open("wb") as dst:
                shutil.copyfileobj(src, dst, length=32 * 1024 * 1024)


def wim_restore_fastest(source_wim: Path, destination: Path) -> None:
    wimlib = require_wimlib()
    temp_extract_root = destination.with_name(destination.name + "__extracting")

    if temp_extract_root.exists():
        shutil.rmtree(temp_extract_root)

    temp_extract_root.mkdir(parents=True, exist_ok=True)

    command = [
        wimlib,
        "apply",
        str(source_wim),
        "1",
        str(temp_extract_root),
        "--no-acls",
    ]

    result = subprocess.run(
        command,
        capture_output=True,
        text=True,
        **subprocess_no_window_kwargs(),
    )

    if result.returncode != 0:
        shutil.rmtree(temp_extract_root, ignore_errors=True)
        output = (result.stdout + result.stderr).strip()
        raise RuntimeError(output or f"wimlib apply failed: {result.returncode}")

    replace_destination_from_extracted_root(temp_extract_root, destination)


def validate_zip_member(member: ZipInfo) -> tuple[str, ...] | None:
    parts = Path(member.filename).parts
    clean_parts = tuple(part for part in parts if part not in ("", "."))

    if not clean_parts:
        return None

    if any(part == ".." for part in clean_parts):
        return None

    if Path(member.filename).is_absolute():
        return None

    return clean_parts


def validate_tar_member(member: tarfile.TarInfo) -> tuple[str, ...] | None:
    parts = Path(member.name).parts
    clean_parts = tuple(part for part in parts if part not in ("", "."))

    if not clean_parts:
        return None

    if any(part == ".." for part in clean_parts):
        return None

    if Path(member.name).is_absolute():
        return None

    if member.issym() or member.islnk() or member.isdev():
        return None

    return clean_parts


def extract_tar_save_contents(source_tar: Path, destination: Path) -> None:
    destination.mkdir(parents=True, exist_ok=True)

    with tarfile.open(source_tar, "r:*") as archive:
        members = archive.getmembers()

        top_level = None
        for member in members:
            parts = validate_tar_member(member)
            if parts:
                top_level = parts[0]
                break

        for member in members:
            parts = validate_tar_member(member)
            if not parts:
                continue

            if top_level and len(parts) > 1 and parts[0] == top_level:
                parts = parts[1:]

            if not parts:
                continue

            target = destination.joinpath(*parts)

            if member.isdir():
                target.mkdir(parents=True, exist_ok=True)
                continue

            if not member.isfile():
                continue

            target.parent.mkdir(parents=True, exist_ok=True)

            source_file = archive.extractfile(member)
            if source_file is None:
                continue

            with source_file, target.open("wb") as dst:
                shutil.copyfileobj(source_file, dst, length=32 * 1024 * 1024)


def restore_from_tar_fastest(source_tar: Path, destination: Path) -> bool:
    tar_exe = shutil.which("tar")
    if not tar_exe:
        return False

    temp_extract_root = destination.with_name(destination.name + "__extracting")

    if temp_extract_root.exists():
        shutil.rmtree(temp_extract_root)

    temp_extract_root.mkdir(parents=True, exist_ok=True)

    command = [
        tar_exe,
        "-xf",
        str(source_tar),
        "-C",
        str(temp_extract_root),
    ]

    result = subprocess.run(
        command,
        capture_output=True,
        text=True,
        **subprocess_no_window_kwargs(),
    )

    if result.returncode != 0:
        shutil.rmtree(temp_extract_root, ignore_errors=True)
        output = (result.stdout + result.stderr).strip()
        raise RuntimeError(output or f"tar extract failed: {result.returncode}")

    children = [item for item in temp_extract_root.iterdir()]
    extracted_root = children[0] if len(children) == 1 and children[0].is_dir() else temp_extract_root
    replace_destination_from_extracted_root(extracted_root, destination)
    return True


class BackupWorker(QThread):
    log = pyqtSignal(str)
    done = pyqtSignal(str)
    failed = pyqtSignal(str)

    def __init__(self, save: SaveEntry, backup_root: Path, method: str):
        super().__init__()
        self.save = save
        self.backup_root = backup_root
        self.method = method

    def run(self) -> None:
        destination: Path | None = None

        try:
            self.backup_root.mkdir(parents=True, exist_ok=True)
            self.log.emit(f"Backing up {self.save.mode}/{self.save.name}")

            started = time.perf_counter()

            if self.method == "zip":
                destination = unique_destination(
                    self.backup_root,
                    self.save.mode,
                    self.save.name,
                    ".zip",
                )
                self.log.emit("Method: zip store")
                zip_store_fast(self.save.path, destination)

            elif self.method == "folder":
                destination = unique_destination(
                    self.backup_root,
                    self.save.mode,
                    self.save.name,
                )
                self.log.emit("Method: folder copy")
                copy_tree_fast(self.save.path, destination)

            elif self.method == "7z":
                destination = unique_destination(
                    self.backup_root,
                    self.save.mode,
                    self.save.name,
                    ".7z",
                )
                self.log.emit("Method: 7-Zip full archive")
                seven_zip_fast_full_backup(self.save.path, destination)

            elif self.method == "wim":
                destination = unique_destination(
                    self.backup_root,
                    self.save.mode,
                    self.save.name,
                    ".wim",
                )
                self.log.emit("Method: WIM full backup")
                wim_full_backup_fastest(self.save.path, destination)

            else:
                raise RuntimeError(f"Unknown backup method: {self.method}")

            elapsed = time.perf_counter() - started

            if destination.is_file():
                size = destination.stat().st_size
                mb = size / 1024 / 1024
                speed = mb / elapsed if elapsed > 0 else 0

                self.log.emit(f"Backup size: {mb:.1f} MB")
                self.log.emit(f"Elapsed: {elapsed:.2f} s")
                self.log.emit(f"Rate: {speed:.1f} MB/s")
            else:
                self.log.emit(f"Elapsed: {elapsed:.2f} s")

            self.done.emit(str(destination))

        except Exception as exc:
            if destination and destination.exists():
                try:
                    if destination.is_dir():
                        shutil.rmtree(destination)
                    else:
                        destination.unlink()
                except OSError:
                    pass

            self.failed.emit(str(exc))


class RestoreWorker(QThread):
    log = pyqtSignal(str)
    done = pyqtSignal(str)
    failed = pyqtSignal(str)

    def __init__(self, backup: BackupEntry, save_root: Path):
        super().__init__()
        self.backup = backup
        self.save_root = save_root

    def run(self) -> None:
        moved_existing: Path | None = None
        destination = self.save_root / self.backup.mode / self.backup.name

        try:
            destination.parent.mkdir(parents=True, exist_ok=True)
            self.log.emit(f"Restoring {self.backup.mode}/{self.backup.name}")

            started = time.perf_counter()

            if destination.exists():
                moved_existing = unique_sidecar_path(destination, "before_restore")
                destination.rename(moved_existing)
                self.log.emit(f"Moved current save aside: {moved_existing}")

            if self.backup.kind == "zip":
                extract_zip_save_contents(self.backup.path, destination)

            elif self.backup.kind == "folder":
                copy_tree_fast(self.backup.path, destination)

            elif self.backup.kind == "7z":
                extract_7z_save_contents(self.backup.path, destination)

            elif self.backup.kind == "wim":
                wim_restore_fastest(self.backup.path, destination)

            elif self.backup.kind == "tar":
                if not restore_from_tar_fastest(self.backup.path, destination):
                    extract_tar_save_contents(self.backup.path, destination)

            elif self.backup.kind == "tar.gz":
                extract_tar_save_contents(self.backup.path, destination)

            else:
                raise RuntimeError(f"Unknown backup kind: {self.backup.kind}")

            elapsed = time.perf_counter() - started
            self.log.emit(f"Restore elapsed: {elapsed:.2f} s")

            self.done.emit(str(destination))

        except Exception as exc:
            if destination.exists():
                try:
                    if destination.is_dir():
                        shutil.rmtree(destination)
                    else:
                        destination.unlink()
                except OSError:
                    pass

            if moved_existing and moved_existing.exists() and not destination.exists():
                try:
                    moved_existing.rename(destination)
                    self.log.emit("Restore failed; original save folder was put back.")
                except OSError:
                    pass

            self.failed.emit(str(exc))


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()

        self.setWindowTitle("Project Zomboid Save Backup")
        self.resize(980, 680)

        self.saves: list[SaveEntry] = []
        self.backups: list[BackupEntry] = []
        self.worker: BackupWorker | RestoreWorker | None = None
        self.status_boxes: list[QMessageBox] = []

        self.save_root = QLineEdit(str(DEFAULT_SAVE_ROOT))
        self.backup_root = QLineEdit(str(DEFAULT_BACKUP_ROOT))

        self.method = QComboBox()
        self.method.addItem("Fast zip, no compression", "zip")
        self.method.addItem("Fast folder copy", "folder")
        self.method.addItem("7-Zip full archive", "7z")
        self.method.addItem("Fast WIM full backup", "wim")

        self.main_menu_confirm = QCheckBox(
            "I have returned to the main menu before backing up or restoring"
        )

        self.refresh_button = QPushButton("Refresh Saves")
        self.refresh_backups_button = QPushButton("Refresh Backups")
        self.backup_button = QPushButton("Backup Selected Save")
        self.restore_button = QPushButton("Restore Selected Backup")
        self.open_backup_button = QPushButton("Open Backup Folder")

        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["Mode", "Save", "Updated", "Path"])
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)

        self.backup_table = QTableWidget(0, 5)
        self.backup_table.setHorizontalHeaderLabels(["Mode", "Save", "Type", "Updated", "Path"])
        self.backup_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.backup_table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.backup_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.backup_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.backup_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.backup_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.backup_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)

        self.progress = QProgressBar()
        self.progress.setRange(0, 1)
        self.progress.setValue(0)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)

        self.build_layout()
        self.bind_events()
        self.refresh_saves()
        self.refresh_backups()

    def build_layout(self) -> None:
        root_group = QGroupBox("Folders")
        grid = QGridLayout(root_group)

        browse_save = QPushButton("Browse")
        browse_backup = QPushButton("Browse")

        browse_save.clicked.connect(self.choose_save_root)
        browse_backup.clicked.connect(self.choose_backup_root)

        grid.addWidget(QLabel("Save root"), 0, 0)
        grid.addWidget(self.save_root, 0, 1)
        grid.addWidget(browse_save, 0, 2)

        grid.addWidget(QLabel("Backup root"), 1, 0)
        grid.addWidget(self.backup_root, 1, 1)
        grid.addWidget(browse_backup, 1, 2)

        grid.addWidget(QLabel("Backup method"), 2, 0)
        grid.addWidget(self.method, 2, 1)

        action_row = QHBoxLayout()
        action_row.addWidget(self.refresh_button)
        action_row.addWidget(self.refresh_backups_button)
        action_row.addWidget(self.backup_button)
        action_row.addWidget(self.restore_button)
        action_row.addWidget(self.open_backup_button)
        action_row.addStretch(1)

        layout = QVBoxLayout()
        layout.addWidget(root_group)
        layout.addWidget(self.main_menu_confirm)
        layout.addLayout(action_row)
        layout.addWidget(QLabel("Saves"))
        layout.addWidget(self.table, 1)
        layout.addWidget(QLabel("Backups"))
        layout.addWidget(self.backup_table, 1)
        layout.addWidget(self.progress)
        layout.addWidget(self.log_box)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def bind_events(self) -> None:
        self.refresh_button.clicked.connect(self.refresh_saves)
        self.refresh_backups_button.clicked.connect(self.refresh_backups)
        self.backup_button.clicked.connect(self.backup_selected)
        self.restore_button.clicked.connect(self.restore_selected_backup)
        self.open_backup_button.clicked.connect(self.open_backup_folder)

    def append_log(self, message: str) -> None:
        self.log_box.append(message)

    def show_status_message(self, title: str, message: str, icon: QMessageBox.Icon) -> None:
        box = QMessageBox(self)
        box.setWindowTitle(title)
        box.setText(message)
        box.setIcon(icon)
        box.setStandardButtons(QMessageBox.StandardButton.Ok)
        box.setModal(False)
        box.finished.connect(
            lambda _result, popup=box: self.status_boxes.remove(popup)
            if popup in self.status_boxes
            else None
        )
        self.status_boxes.append(box)
        box.open()
        QTimer.singleShot(10000, box.close)

    def choose_save_root(self) -> None:
        path = QFileDialog.getExistingDirectory(
            self,
            "Select Zomboid Saves Folder",
            self.save_root.text(),
        )

        if path:
            self.save_root.setText(path)
            self.refresh_saves()

    def choose_backup_root(self) -> None:
        path = QFileDialog.getExistingDirectory(
            self,
            "Select Backup Folder",
            self.backup_root.text(),
        )

        if path:
            self.backup_root.setText(path)
            self.refresh_backups()

    def refresh_saves(self) -> None:
        root = Path(self.save_root.text()).expanduser()
        self.saves = scan_saves(root)

        self.table.setRowCount(len(self.saves))

        for row, save in enumerate(self.saves):
            values = [
                save.mode,
                save.name,
                format_time(save.updated),
                str(save.path),
            ]

            for col, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.table.setItem(row, col, item)

        if self.saves:
            self.table.selectRow(0)

        self.append_log(f"Found {len(self.saves)} saves in {root}")

    def refresh_backups(self) -> None:
        root = Path(self.backup_root.text()).expanduser()
        self.backups = scan_backups(root)

        self.backup_table.setRowCount(len(self.backups))

        for row, backup in enumerate(self.backups):
            values = [
                backup.mode,
                backup.name,
                backup.kind,
                format_time(backup.updated),
                str(backup.path),
            ]

            for col, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.backup_table.setItem(row, col, item)

        if self.backups:
            self.backup_table.selectRow(0)

        self.append_log(f"Found {len(self.backups)} backups in {root}")

    def selected_save(self) -> SaveEntry | None:
        row = self.table.currentRow()
        if row < 0 or row >= len(self.saves):
            return None
        return self.saves[row]

    def selected_backup(self) -> BackupEntry | None:
        row = self.backup_table.currentRow()
        if row < 0 or row >= len(self.backups):
            return None
        return self.backups[row]

    def backup_selected(self) -> None:
        if self.worker and self.worker.isRunning():
            return

        save = self.selected_save()
        if not save:
            QMessageBox.warning(self, "No save selected", "Select a save first.")
            return

        if not self.main_menu_confirm.isChecked():
            QMessageBox.warning(
                self,
                "Main menu required",
                "Project Zomboid does not support safe hot backup. Return to the main menu first, then tick the checkbox.",
            )
            return

        backup_root = Path(self.backup_root.text()).expanduser()
        method = self.method.currentData()

        self.set_busy(True)
        self.progress.setRange(0, 0)

        self.worker = BackupWorker(save, backup_root, method)
        self.worker.log.connect(self.append_log)
        self.worker.done.connect(self.backup_done)
        self.worker.failed.connect(self.backup_failed)
        self.worker.start()

    def restore_selected_backup(self) -> None:
        if self.worker and self.worker.isRunning():
            return

        backup = self.selected_backup()
        if not backup:
            QMessageBox.warning(self, "No backup selected", "Select a backup first.")
            return

        if not self.main_menu_confirm.isChecked():
            QMessageBox.warning(
                self,
                "Main menu required",
                "Project Zomboid does not support safe hot restore. Return to the main menu first, then tick the checkbox.",
            )
            return

        save_root = Path(self.save_root.text()).expanduser()
        destination = save_root / backup.mode / backup.name

        self.append_log(f"Restore target: {destination}")
        self.append_log("Existing save folder will be renamed before restore.")

        self.set_busy(True)
        self.progress.setRange(0, 0)

        self.worker = RestoreWorker(backup, save_root)
        self.worker.log.connect(self.append_log)
        self.worker.done.connect(self.restore_done)
        self.worker.failed.connect(self.restore_failed)
        self.worker.start()

    def set_busy(self, busy: bool) -> None:
        self.backup_button.setEnabled(not busy)
        self.restore_button.setEnabled(not busy)
        self.refresh_button.setEnabled(not busy)
        self.refresh_backups_button.setEnabled(not busy)
        self.method.setEnabled(not busy)

    def backup_done(self, destination: str) -> None:
        self.progress.setRange(0, 1)
        self.progress.setValue(1)
        self.set_busy(False)
        self.append_log(f"Backup complete: {destination}")
        self.refresh_backups()
        QMessageBox.information(self, "Backup complete", destination)

    def backup_failed(self, error: str) -> None:
        self.progress.setRange(0, 1)
        self.progress.setValue(0)
        self.set_busy(False)
        self.append_log(f"Backup failed: {error}")
        QMessageBox.critical(self, "Backup failed", error)

    def restore_done(self, destination: str) -> None:
        self.progress.setRange(0, 1)
        self.progress.setValue(1)
        self.set_busy(False)
        self.append_log(f"Restore complete: {destination}")
        self.refresh_saves()
        self.show_status_message(
            "Restore complete",
            destination,
            QMessageBox.Icon.Information,
        )

    def restore_failed(self, error: str) -> None:
        self.progress.setRange(0, 1)
        self.progress.setValue(0)
        self.set_busy(False)
        self.append_log(f"Restore failed: {error}")
        self.show_status_message(
            "Restore failed",
            error,
            QMessageBox.Icon.Critical,
        )

    def open_backup_folder(self) -> None:
        path = Path(self.backup_root.text()).expanduser()
        path.mkdir(parents=True, exist_ok=True)

        if os.name == "nt":
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])


def main() -> int:
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
