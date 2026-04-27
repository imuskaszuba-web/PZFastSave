import os
import shutil
import subprocess
import sys
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


APP_DIR = Path(__file__).resolve().parent
DEFAULT_SAVE_ROOT = Path.home() / "Zomboid" / "Saves"
DEFAULT_BACKUP_ROOT = Path.home() / "Zomboid" / "SaveBackups"


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


def parse_backup_name(path: Path) -> tuple[str, str] | None:
    stem = path.stem if path.is_file() and path.suffix.lower() == ".zip" else path.name
    parts = stem.split("__", 2)
    if len(parts) != 3:
        return None
    return parts[1], parts[2]


def scan_backups(backup_root: Path) -> list[BackupEntry]:
    if not backup_root.exists():
        return []

    backups: list[BackupEntry] = []
    for path in backup_root.iterdir():
        if not path.is_dir() and not (path.is_file() and path.suffix.lower() == ".zip"):
            continue

        parsed = parse_backup_name(path)
        if not parsed:
            continue

        try:
            updated = path.stat().st_mtime
        except OSError:
            updated = 0

        mode, name = parsed
        kind = "folder" if path.is_dir() else "zip"
        backups.append(BackupEntry(mode, name, path, updated, kind))

    backups.sort(key=lambda item: item.updated, reverse=True)
    return backups


def unique_destination(root: Path, mode: str, name: str, suffix: str = "") -> Path:
    stamp = time.strftime("%Y%m%d-%H%M%S")
    base = f"{stamp}__{safe_name(mode)}__{safe_name(name)}{suffix}"
    candidate = root / base
    counter = 2
    while candidate.exists():
        candidate = root / f"{base}_{counter}"
        counter += 1
    return candidate


def run_robocopy(source: Path, destination: Path) -> None:
    destination.mkdir(parents=True, exist_ok=True)
    command = [
        "robocopy",
        str(source),
        str(destination),
        "/E",
        "/MT:16",
        "/R:1",
        "/W:1",
        "/NFL",
        "/NDL",
        "/NJH",
        "/NJS",
        "/NP",
    ]
    result = subprocess.run(command, capture_output=True, text=True)
    if result.returncode >= 8:
        raise RuntimeError((result.stdout + result.stderr).strip() or f"robocopy failed: {result.returncode}")


def copy_tree_fast(source: Path, destination: Path) -> None:
    if os.name == "nt" and shutil.which("robocopy"):
        run_robocopy(source, destination)
        return

    shutil.copytree(source, destination, copy_function=shutil.copy2)


def zip_store_fast(source: Path, destination_zip: Path) -> None:
    with zipfile.ZipFile(destination_zip, "w", compression=zipfile.ZIP_STORED, allowZip64=True) as archive:
        for file_path in source.rglob("*"):
            if not file_path.is_file():
                continue
            archive.write(file_path, file_path.relative_to(source.parent))


def unique_sidecar_path(path: Path, marker: str) -> Path:
    stamp = time.strftime("%Y%m%d-%H%M%S")
    candidate = path.with_name(f"{path.name}__{marker}__{stamp}")
    counter = 2
    while candidate.exists():
        candidate = path.with_name(f"{path.name}__{marker}__{stamp}_{counter}")
        counter += 1
    return candidate


def validate_zip_member(member: ZipInfo) -> tuple[str, ...] | None:
    parts = Path(member.filename).parts
    clean_parts = tuple(part for part in parts if part not in ("", "."))
    if not clean_parts or any(part == ".." for part in clean_parts):
        return None
    return clean_parts


def extract_zip_save_contents(source_zip: Path, destination: Path) -> None:
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
                shutil.copyfileobj(src, dst, length=1024 * 1024)


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
        try:
            self.backup_root.mkdir(parents=True, exist_ok=True)
            self.log.emit(f"Backing up {self.save.mode}/{self.save.name}")

            if self.method == "folder":
                destination = unique_destination(self.backup_root, self.save.mode, self.save.name)
                copy_tree_fast(self.save.path, destination)
            else:
                destination = unique_destination(self.backup_root, self.save.mode, self.save.name, ".zip")
                zip_store_fast(self.save.path, destination)

            self.done.emit(str(destination))
        except Exception as exc:
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

            if destination.exists():
                moved_existing = unique_sidecar_path(destination, "before_restore")
                destination.rename(moved_existing)
                self.log.emit(f"Moved current save aside: {moved_existing}")

            if self.backup.kind == "folder":
                copy_tree_fast(self.backup.path, destination)
            else:
                extract_zip_save_contents(self.backup.path, destination)

            self.done.emit(str(destination))
        except Exception as exc:
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
        self.method.addItem("Fast folder copy (recommended)", "folder")
        self.method.addItem("Fast zip, no compression", "zip")

        self.main_menu_confirm = QCheckBox("I have returned to the main menu before backing up or restoring")
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
        box.finished.connect(lambda _result, popup=box: self.status_boxes.remove(popup) if popup in self.status_boxes else None)
        self.status_boxes.append(box)
        box.open()
        QTimer.singleShot(10000, box.close)

    def choose_save_root(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Select Zomboid Saves Folder", self.save_root.text())
        if path:
            self.save_root.setText(path)
            self.refresh_saves()

    def choose_backup_root(self) -> None:
        path = QFileDialog.getExistingDirectory(self, "Select Backup Folder", self.backup_root.text())
        if path:
            self.backup_root.setText(path)
            self.refresh_backups()

    def refresh_saves(self) -> None:
        root = Path(self.save_root.text()).expanduser()
        self.saves = scan_saves(root)
        self.table.setRowCount(len(self.saves))
        for row, save in enumerate(self.saves):
            values = [save.mode, save.name, format_time(save.updated), str(save.path)]
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
            values = [backup.mode, backup.name, backup.kind, format_time(backup.updated), str(backup.path)]
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
        self.append_log("Existing save folders are renamed as before_restore safety copies before restore.")

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
        self.show_status_message("Restore complete", destination, QMessageBox.Icon.Information)

    def restore_failed(self, error: str) -> None:
        self.progress.setRange(0, 1)
        self.progress.setValue(0)
        self.set_busy(False)
        self.append_log(f"Restore failed: {error}")
        self.show_status_message("Restore failed", error, QMessageBox.Icon.Critical)

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
