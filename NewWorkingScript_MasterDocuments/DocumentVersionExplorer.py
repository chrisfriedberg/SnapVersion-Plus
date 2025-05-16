# -*- coding: utf-8 -*-
import os
import sys
import json
import subprocess
import re
import difflib
import ctypes
from datetime import datetime
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QTreeWidget, QTreeWidgetItem,
    QPushButton, QHBoxLayout, QVBoxLayout, QWidget, QMessageBox, QDialog,
    QTextEdit, QDialogButtonBox, QMenuBar, QSplitter, QLabel, QSystemTrayIcon,
    QMenu, QHeaderView, QSizePolicy
)
from PyQt6.QtGui import QAction, QIcon
from PyQt6.QtCore import Qt, QPoint

# Determine the base directory dynamically.
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Constants
APP_NAME = "SnapVersion+Plus - DocVExplorer"
try:
    SCRIPT_NAME = os.path.splitext(os.path.basename(__file__))[0]
except NameError:
    SCRIPT_NAME = "DocumentVersionExplorer"

LOG_CONFIG_FILE = os.path.join(BASE_DIR, f"{SCRIPT_NAME}_log.json")
DEFAULT_ACTION_LOG_FILE = os.path.join(BASE_DIR, f"{SCRIPT_NAME}_default_log.txt")
NOTEPADPP_PATH = r"C:\Program Files\Notepad++\notepad++.exe"

# --- Utility Functions ---
def log_action_global(log_file_path, message):
    """Log an action with duplicate prevention within 1 second."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"[{timestamp}] {message}\n"
    try:
        os.makedirs(os.path.dirname(log_file_path), exist_ok=True)
        if os.path.exists(log_file_path):
            with open(log_file_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
                if lines and lines[-1].strip() == log_entry.strip():
                    return
        with open(log_file_path, 'a', encoding='utf-8') as f:
            f.write(log_entry)
            f.flush()
    except Exception as e:
        print(f"!!! GLOBAL LOGGING ERROR to {log_file_path}: {str(e)} !!!\nOriginal message: {log_entry.strip()}")

# --- Custom Tree Widget Items for Sorting ---
class SortableTreeWidgetItem(QTreeWidgetItem):
    def __lt__(self, other):
        column = self.treeWidget().sortColumn()
        if column == 0:
            my_data = self.data(column, Qt.ItemDataRole.UserRole)
            other_data = other.data(column, Qt.ItemDataRole.UserRole)
            if my_data is None:
                return True
            if other_data is None:
                return False
            try:
                return float(my_data) < float(other_data)
            except (ValueError, TypeError):
                return self.text(column).lower() < other.text(column).lower()
        elif column == 2:
            try:
                return int(self.text(column)) < int(other.text(column))
            except ValueError:
                return self.text(column).lower() < other.text(column).lower()
        else:
            return self.text(column).lower() < other.text(column).lower()

class BackupTreeWidgetItem(QTreeWidgetItem):
    """Custom item for the backup tree to enable proper sorting on Date/Time column. Custom."""
    def __lt__(self, other):
        column = self.treeWidget().sortColumn()
        if column == 0:  # Date/Time column
            my_data = self.data(column, Qt.ItemDataRole.UserRole)
            other_data = other.data(column, Qt.ItemDataRole.UserRole)
            if my_data is None:
                return True
            if other_data is None:
                return False
            try:
                return float(my_data) < float(other_data)
            except (ValueError, TypeError):
                return self.text(column).lower() < other.text(column).lower()
        else:
            return self.text(column).lower() < other.text(column).lower()

# --- Dialog Classes ---
class PreviewDialog(QDialog):
    def __init__(self, file_path, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Preview File")
        self.resize(600, 400)
        layout = QVBoxLayout()
        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)
        try:
            with open(file_path, 'r', encoding='utf-8-sig') as f:
                self.text_edit.setText(f.read())
        except Exception as e:
            self.text_edit.setText(f"Error reading file: {str(e)}")
            log_action_global(DEFAULT_ACTION_LOG_FILE, f"ERROR: Preview read {file_path}: {str(e)}")
        layout.addWidget(self.text_edit)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.setLayout(layout)

class MetaTagEditor(QDialog):
    def __init__(self, file_path, current_meta, parent_logger, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Meta Tag")
        self.resize(400, 150)
        self.file_path = file_path
        self.logger = parent_logger
        layout = QVBoxLayout()
        self.text_edit = QTextEdit()
        self.text_edit.setPlainText(current_meta)
        layout.addWidget(QLabel(f"Editing metadata for:\n{os.path.basename(file_path)}"))
        layout.addWidget(self.text_edit)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.save_meta)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.setLayout(layout)

    def save_meta(self):
        try:
            content = self.text_edit.toPlainText().strip()
            with open(f"{self.file_path}:source", "w", encoding="utf-8") as f:
                f.write(content)
            try:
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                history_entry = f"[{timestamp}] {content}"
                with open(f"{self.file_path}:meta_audit", "a", encoding="utf-8") as f:
                    f.write(history_entry + "\n")
            except Exception as e:
                self.logger(f"ERROR: Audit write fail {self.file_path}: {e}")
        except Exception as e:
            self.logger(f"ERROR: Meta update fail {self.file_path}: {e}")
        self.accept()

class MetadataHistoryDialog(QDialog):
    def __init__(self, file_path, parent_logger, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Metadata History")
        self.resize(400, 300)
        self.logger = parent_logger
        layout = QVBoxLayout()
        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)
        history = self._get_log_history_for_file(file_path)
        if history:
            self.text_edit.setPlainText("\n".join(history))
        else:
            self.text_edit.setText("No metadata history available.")
        layout.addWidget(QLabel(f"Metadata history for:\n{os.path.basename(file_path)}"))
        layout.addWidget(self.text_edit)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.setLayout(layout)

    def _get_log_history_for_file(self, file_path):
        """Load metadata history, sorted newest-first, with robust deduplication."""
        try:
            ads_path = f"{file_path}:meta_audit"
            if not os.path.exists(ads_path):
                if not os.path.exists(file_path):
                    self.logger(f"WARNING: Main file not found for history: {file_path}")
                return []
            with open(ads_path, "r", encoding="utf-8") as f:
                raw_history = [line.strip() for line in f.readlines() if line.strip()]
            seen = set()
            unique_history = []
            for entry in raw_history[::-1]:
                if entry not in seen:
                    unique_history.append(entry)
                    seen.add(entry)
            unique_history.reverse()
            if len(raw_history) != len(unique_history):
                self.logger(f"INFO: Deduplicated history for {os.path.basename(file_path)} ({len(raw_history)} -> {len(unique_history)} entries)")
            return unique_history
        except FileNotFoundError:
            return []
        except Exception as e:
            self.logger(f"ERROR: Read history fail {file_path}: {e}")
            return []

# --- Main Application Class ---
class DocumentVersionExplorer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(1000, 600)
        self.production_directory = ''
        self.backup_directory = ''
        self.action_log_file = DEFAULT_ACTION_LOG_FILE
        self.log_action("INFO: Application started")
        log_data = self.load_config()
        self.production_directory = log_data.get('production_directory', '')
        self.backup_directory = log_data.get('backup_directory', '')
        self.main_icon_path = log_data.get('main_icon_path', '')
        self.tray_icon_path = log_data.get('tray_icon_path', '')
        self.current_document = None
        self.setup_tray_icon()
        self.setup_main_icon()
        self.setup_menu_bar()
        self.setup_main_ui_original()
        if self.production_directory and self.backup_directory:
            self.load_master_documents()

    def log_action(self, message):
        log_action_global(self.action_log_file, message)

    def load_config(self):
        if not os.path.exists(LOG_CONFIG_FILE):
            self.log_action(f"INFO: Config missing {LOG_CONFIG_FILE}")
            return {}
        try:
            with open(LOG_CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except json.JSONDecodeError as e:
            self.log_action(f"ERROR: Corrupt config {LOG_CONFIG_FILE}: {str(e)}")
            QMessageBox.warning(self, "Config Error", f"Config corrupt. Using defaults.")
            try:
                os.rename(LOG_CONFIG_FILE, f"{LOG_CONFIG_FILE}.corrupt_{datetime.now():%Y%m%d%H%M%S}")
            except OSError as re:
                self.log_action(f"ERROR: Rename corrupt config fail: {re}")
            return {}
        except Exception as e:
            self.log_action(f"ERROR: Load config fail {LOG_CONFIG_FILE}: {str(e)}")
            return {}

    def save_config(self, key, value):
        data = self.load_config()
        data[key] = value
        try:
            os.makedirs(os.path.dirname(LOG_CONFIG_FILE), exist_ok=True)
            with open(LOG_CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4)
            self.log_action(f"INFO: Saved '{key}' to config: {value}")
        except Exception as e:
            self.log_action(f"ERROR: Save config fail {LOG_CONFIG_FILE}: {str(e)}")
            QMessageBox.warning(self, "Config Error", f"Failed save '{key}'.")

    def setup_tray_icon(self):
        self.tray_icon = QSystemTrayIcon(self)
        try:
            icon = QIcon.fromTheme("application-x-executable")
            if self.tray_icon_path and os.path.exists(self.tray_icon_path):
                loaded_icon = QIcon(self.tray_icon_path)
                if not loaded_icon.isNull():
                    icon = loaded_icon
                else:
                    self.log_action(f"WARNING: Tray icon load fail: {self.tray_icon_path}")
            self.tray_icon.setIcon(icon)
        except Exception as e:
            self.log_action(f"ERROR: Tray icon set fail: {str(e)}")
            self.tray_icon.setIcon(QIcon.fromTheme("application-x-executable"))
        self.tray_icon.setToolTip(APP_NAME)
        tray_menu = QMenu()
        restore_action = QAction("Restore", self)
        restore_action.triggered.connect(self.restore_from_tray)
        tray_menu.addAction(restore_action)
        quit_action = QAction("Quit", self)
        quit_action.triggered.connect(self.quit_application)
        tray_menu.addAction(quit_action)
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()

    def setup_main_icon(self):
        try:
            icon = QIcon.fromTheme("application-x-executable")
            if self.main_icon_path and os.path.exists(self.main_icon_path):
                loaded_icon = QIcon(self.main_icon_path)
                if not loaded_icon.isNull():
                    icon = loaded_icon
                else:
                    self.log_action(f"WARNING: Main icon load fail: {self.main_icon_path}")
            self.setWindowIcon(icon)
        except Exception as e:
            self.log_action(f"ERROR: Main icon set fail: {str(e)}")
            self.setWindowIcon(QIcon.fromTheme("application-x-executable"))

    def setup_menu_bar(self):
        self.menu_bar = QMenuBar(self)
        self.setMenuBar(self.menu_bar)
        file_menu = self.menu_bar.addMenu("File")
        open_prod_dir_action = QAction("Set Production Directory", self)
        open_prod_dir_action.triggered.connect(self.set_production_directory)
        file_menu.addAction(open_prod_dir_action)
        open_backup_dir_action = QAction("Set Backup Directory", self)
        open_backup_dir_action.triggered.connect(self.set_backup_directory)
        file_menu.addAction(open_backup_dir_action)
        file_menu.addSeparator()
        open_log_action = QAction("Open Log File", self)
        open_log_action.triggered.connect(self.open_log_file)
        file_menu.addAction(open_log_action)
        file_menu.addSeparator()
        close_action = QAction("Close", self)
        close_action.triggered.connect(self.quit_application)
        file_menu.addAction(close_action)
        edit_menu = self.menu_bar.addMenu("Edit")
        select_main_icon_action = QAction("Select Main Bar Icon", self)
        select_main_icon_action.setToolTip("Recommended: 16x16 or 32x32 PNG or ICO file")
        select_main_icon_action.triggered.connect(self.select_main_icon)
        edit_menu.addAction(select_main_icon_action)
        select_tray_icon_action = QAction("Select Tray Icon", self)
        select_tray_icon_action.setToolTip("Recommended: 16x16 or 32x32 PNG or ICO file")
        select_tray_icon_action.triggered.connect(self.select_tray_icon)
        edit_menu.addAction(select_tray_icon_action)
        self.menu_bar.addMenu("View")
        self.menu_bar.addMenu("Help")

    def setup_main_ui_original(self):
        """Setup the main splitter UI with fixed path label display."""
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QHBoxLayout(self.central_widget)
        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        self.splitter.setHandleWidth(5)
        self.splitter.setStyleSheet("QSplitter::handle { background-color: #a83232; } QSplitter::handle:hover { background-color: #8B0000; }")
        self.main_layout.addWidget(self.splitter)

        # Left Pane
        self.left_widget = QWidget()
        self.left_widget.setFixedWidth(740)
        self.left_layout = QVBoxLayout(self.left_widget)
        self.prod_dir_title = QLabel("Main File Directory:")
        self.prod_dir_title.setStyleSheet("font-weight: bold; color: white;")
        self.left_layout.addWidget(self.prod_dir_title)
        self.prod_dir_container = QHBoxLayout()
        self.explore_btn = QPushButton("Explore")
        self.explore_btn.setStyleSheet("background-color: #a83232; color: white; font-weight: bold;")
        self.explore_btn.setFixedHeight(50)
        self.explore_btn.setFixedWidth(140)
        self.explore_btn.clicked.connect(self.explore_production_directory)
        self.prod_dir_container.addWidget(self.explore_btn)
        self.prod_dir_label = QLabel(self.production_directory or 'Not set')
        self.prod_dir_label.setStyleSheet("background-color: #333; color: white; padding: 5px; border: 1px solid #0078D7;")
        self.prod_dir_label.setWordWrap(True)
        self.prod_dir_label.setFixedHeight(50)
        self.prod_dir_label.setMaximumWidth(580)
        self.prod_dir_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.prod_dir_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.prod_dir_container.addWidget(self.prod_dir_label)
        self.left_layout.addLayout(self.prod_dir_container)
        self.master_list = QTreeWidget()
        self.master_list.setMinimumWidth(600)
        self.master_list.setMaximumWidth(600)
        self.master_list.setMinimumHeight(400)
        self.master_list.setColumnCount(3)
        self.master_list.setHeaderLabels(["Last Modified", "File Name", "Backups"])
        self.master_list.itemClicked.connect(self.load_versions)
        self.master_list.itemSelectionChanged.connect(self.update_tray_tooltip)
        self.master_list.setStyleSheet("QTreeWidget { padding: 5px; } QTreeWidget::item { padding: 5px; }")
        self.master_list.setAlternatingRowColors(True)
        self.master_list.setSortingEnabled(True)
        self.left_layout.addWidget(self.master_list)
        self.splitter.addWidget(self.left_widget)

        # Right Pane
        self.right_widget = QWidget()
        self.right_widget.setFixedWidth(600)
        self.right_layout = QVBoxLayout(self.right_widget)
        self.backup_dir_title = QLabel("Backup Directory:")
        self.backup_dir_title.setStyleSheet("font-weight: bold; color: white;")
        self.right_layout.addWidget(self.backup_dir_title)
        self.backup_dir_label = QLabel(self.backup_directory or 'Not set')
        self.backup_dir_label.setStyleSheet("background-color: #333; color: white; padding: 5px; border: 1px solid #0078D7;")
        # Fix: Align with left pane's prod_dir_label
        self.backup_dir_label.setWordWrap(True)
        self.backup_dir_label.setFixedHeight(50)
        self.backup_dir_label.setMaximumWidth(580)  # Reduced to match prod_dir_label's effective width
        self.backup_dir_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.backup_dir_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.right_layout.addWidget(self.backup_dir_label)
        self.preview_btn = QPushButton("Preview")
        self.preview_btn.setStyleSheet("background-color: #a83232; color: white; font-weight: bold;")
        self.preview_btn.setFixedWidth(600)
        self.preview_btn.clicked.connect(self.preview_selected)
        self.preview_btn.setEnabled(False)
        self.right_layout.addWidget(self.preview_btn)
        self.tree = QTreeWidget()
        self.tree.setMinimumWidth(600)
        self.tree.setColumnCount(6)
        self.tree.setHeaderLabels(["Date/Time", "Group Alias", "Version", "Changes", "Total Lines", "Meta Tag"])
        self.tree.setStyleSheet("QTreeWidget { padding: 5px; } QTreeWidget::item { padding: 5px; }")
        self.tree.setAlternatingRowColors(True)
        self.tree.itemDoubleClicked.connect(self.handle_tree_double_click)
        self.tree.itemSelectionChanged.connect(self.update_preview_button)
        self.tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self.show_context_menu)
        self.tree.setSortingEnabled(True)
        self.tree.sortByColumn(0, Qt.SortOrder.DescendingOrder)
        self.right_layout.addWidget(self.tree)
        self.button_container = QHBoxLayout()
        self.refresh_files_btn = QPushButton("Refresh Files")
        self.refresh_files_btn.setStyleSheet("background-color: #0078D7; color: white; font-weight: bold;")
        self.refresh_files_btn.clicked.connect(self.refresh_files)
        self.button_container.addWidget(self.refresh_files_btn)
        self.refresh_meta_btn = QPushButton("Refresh Meta Tags")
        self.refresh_meta_btn.setStyleSheet("background-color: #0078D7; color: white; font-weight: bold;")
        self.refresh_meta_btn.clicked.connect(self.refresh_meta_tags)
        self.button_container.addWidget(self.refresh_meta_btn)
        self.close_minimize_container = QHBoxLayout()
        self.close_minimize_container.addStretch()
        self.minimize_btn = QPushButton("Minimize")
        self.minimize_btn.setStyleSheet("background-color: #0078D7; color: white; font-weight: bold;")
        self.minimize_btn.clicked.connect(self.minimize_to_tray)
        self.close_minimize_container.addWidget(self.minimize_btn)
        self.close_btn = QPushButton("Close")
        self.close_btn.setStyleSheet("background-color: #a83232; color: white; font-weight: bold;")
        self.close_btn.clicked.connect(self.quit_application)
        self.close_minimize_container.addWidget(self.close_btn)
        self.bottom_right_layout = QVBoxLayout()
        self.bottom_right_layout.addLayout(self.button_container)
        self.bottom_right_layout.addLayout(self.close_minimize_container)
        self.right_layout.addLayout(self.bottom_right_layout)
        self.splitter.addWidget(self.right_widget)
        self.main_layout.addWidget(self.splitter)

    def _get_base_name(self, filename):
        pattern = r'^(.*)\.\d{4}-\d{2}-\d{2}_\d{6}\.bak$'
        match = re.match(pattern, filename)
        if match:
            return match.group(1)
        return filename.split('.')[0]

    def _get_backup_count(self, base_name):
        if not self.backup_directory or not os.path.isdir(self.backup_directory):
            return 0
        try:
            backups = [f for f in os.listdir(self.backup_directory) if f.startswith(base_name) and f.endswith('.bak')]
            return len(backups)
        except Exception as e:
            self.log_action(f"ERROR: Count backups fail '{base_name}': {str(e)}")
            return 0

    def load_master_documents(self):
        self.master_list.clear()
        if not self.production_directory or not os.path.isdir(self.production_directory):
            self.log_action("ERROR: Prod dir invalid")
            return
        if not self.backup_directory or not os.path.isdir(self.backup_directory):
            self.log_action("ERROR: Backup dir invalid")
            return
        self.log_action(f"INFO: Loading master documents from {self.production_directory}")
        try:
            files_loaded_count = 0
            files = [f for f in os.listdir(self.production_directory) if os.path.isfile(os.path.join(self.production_directory, f))]
            for f_name in files:
                full_path = os.path.join(self.production_directory, f_name)
                try:
                    mtime_float = os.path.getmtime(full_path)
                    mod_time_str = datetime.fromtimestamp(mtime_float).strftime("%Y-%m-%d %H:%M:%S")
                except OSError as e:
                    self.log_action(f"ERROR: Cannot get mod time for {f_name}: {e}")
                    mtime_float = 0.0
                    mod_time_str = "N/A"
                base_name = self._get_base_name(f_name)
                backup_count = self._get_backup_count(base_name)
                item_data = [mod_time_str, f_name, str(backup_count)]
                item = SortableTreeWidgetItem(item_data)
                item.setData(0, Qt.ItemDataRole.UserRole, mtime_float)
                item.setData(1, Qt.ItemDataRole.UserRole, f_name)
                self.master_list.addTopLevelItem(item)
                files_loaded_count += 1
            self.log_action(f"INFO: Loaded {files_loaded_count} master documents")
            self.master_list.setColumnWidth(0, 140)
            self.master_list.setColumnWidth(1, 380)
            self.master_list.setColumnWidth(2, 80)
            self.master_list.sortByColumn(0, Qt.SortOrder.DescendingOrder)
        except Exception as e:
            self.log_action(f"ERROR: Load master docs fail: {str(e)}")
            QMessageBox.warning(self, "Error", "Failed load master docs.")

    def load_versions(self, item):
        filename = item.data(1, Qt.ItemDataRole.UserRole)
        if filename is None:
            self.log_action("ERROR: No filename associated with selected master item.")
            return
        base_name = self._get_base_name(filename)
        self.log_action(f"INFO: Loading versions for: {filename} (Base: {base_name})")
        self.tree.clear()
        if not self.backup_directory or not os.path.isdir(self.backup_directory):
            self.log_action("ERROR: Backup dir invalid.")
            QMessageBox.warning(self, "Error", "Backup directory invalid.")
            return
        try:
            backups = [f for f in os.listdir(self.backup_directory) if f.startswith(base_name) and f.endswith('.bak')]
            self.log_action(f"INFO: Found {len(backups)} backups for base '{base_name}'.")
        except Exception as e:
            self.log_action(f"ERROR: List backups fail: {str(e)}")
            return
        if not backups:
            self.log_action(f"INFO: No backups found for: {base_name}")
            return
        entries = []
        for name in backups:
            path = os.path.join(self.backup_directory, name)
            try:
                timestamp = datetime.fromtimestamp(os.path.getctime(path))
                dt_str = timestamp.strftime("%a %m/%d/%Y %I:%M%p").lower()
                entries.append((timestamp, dt_str, base_name, name, path, timestamp.timestamp()))
            except OSError as e:
                self.log_action(f"ERROR: Access fail {path}: {str(e)}")
                continue
        entries.sort(key=lambda x: x[0], reverse=True)
        diff_counts = []
        total_lines = []
        for i in range(len(entries)):
            curr_path = entries[i][4]
            try:
                with open(curr_path, 'r', encoding='utf-8-sig') as f:
                    line_count = len(f.readlines())
                total_lines.append(str(line_count))
            except Exception as e:
                self.log_action(f"ERROR: Count lines fail {curr_path}: {str(e)}")
                total_lines.append("Error")
            if i == len(entries) - 1:
                diff_counts.append("N/A")
            else:
                next_path = entries[i+1][4]
                try:
                    with open(curr_path, 'r', encoding='utf-8-sig') as f_curr, open(next_path, 'r', encoding='utf-8-sig') as f_next:
                        curr_lines = f_curr.readlines()
                        next_lines = f_next.readlines()
                        diff = list(difflib.unified_diff(next_lines, curr_lines, lineterm=''))
                        change_count = sum(1 for line in diff if line.startswith(('+', '-')) and not line.startswith(('+++', '---')))
                        curr_total = len(curr_lines)
                        next_total = len(next_lines)
                        change_sign = "+" if curr_total > next_total else "-" if curr_total < next_total else ""
                        diff_counts.append(f"{change_sign}{change_count} lines")
                except Exception as e:
                    self.log_action(f"ERROR: Compare fail {curr_path} / {next_path}: {str(e)}")
                    diff_counts.append("Error")
        total_versions = len(entries)
        for idx, (entry_data, diff_count, total_line_count) in enumerate(zip(entries, diff_counts, total_lines)):
            version_number = total_versions - idx
            _, dt, base, name, full_path, timestamp_float = entry_data
            meta = self.read_ads_metadata(full_path)
            item = BackupTreeWidgetItem([dt, base, f"V{version_number}", diff_count, total_line_count, meta])
            item.setData(0, Qt.ItemDataRole.UserRole, timestamp_float)  # Store float for sorting
            item.setData(0, Qt.ItemDataRole.UserRole + 1, full_path)  # Store path in a different role
            self.tree.addTopLevelItem(item)
        self.tree.sortByColumn(0, Qt.SortOrder.DescendingOrder)
        try:
            for i in range(self.tree.columnCount()):
                self.tree.resizeColumnToContents(i)
        except Exception as e:
            self.log_action(f"ERROR: Resize tree columns fail: {str(e)}")

    def read_ads_metadata(self, file_path):
        try:
            ads_path = f"{file_path}:source"
            if not os.path.exists(ads_path):
                if not os.path.exists(file_path):
                    self.log_action(f"WARNING: Main file missing {file_path}")
                return ""
            with open(ads_path, "r", encoding="utf-8") as f:
                return f.read().strip()
        except FileNotFoundError:
            return ""
        except Exception as e:
            self.log_action(f"ERROR: Read ADS source fail {file_path}: {str(e)}")
            return ""

    def read_meta_audit(self, file_path):
        max_retries = 3
        for attempt in range(1, max_retries + 1):
            try:
                ads_path = f"{file_path}:meta_audit"
                if not os.path.exists(ads_path):
                    if not os.path.exists(file_path):
                        pass
                    return []
                with open(ads_path, "r", encoding="utf-8") as f:
                    lines = f.readlines()
                return [line.strip() for line in lines if line.strip()]
            except FileNotFoundError:
                return []
            except Exception as e:
                self.log_action(f"ERROR: Read meta audit fail {file_path} (Att {attempt}): {str(e)}")
                if attempt == max_retries:
                    self.log_action(f"ERROR: Max retries read meta audit {file_path}")
                    return []
        return []

    def append_meta_audit(self, file_path, entries_to_append):
        if not entries_to_append:
            return
        existing_entries_set = set(self.read_meta_audit(file_path))
        new_unique_entries = [entry for entry in entries_to_append if entry.strip() and entry.strip() not in existing_entries_set]
        if not new_unique_entries:
            return
        max_retries = 3
        for attempt in range(1, max_retries + 1):
            try:
                ads_path = f"{file_path}:meta_audit"
                if not os.path.exists(file_path):
                    self.log_action(f"ERROR: Append meta audit fail, file missing: {file_path}")
                    return
                with open(ads_path, "a", encoding="utf-8") as f:
                    f.writelines(f"{entry}\n" for entry in new_unique_entries)
                return
            except Exception as e:
                self.log_action(f"ERROR: Append meta audit fail {file_path} (Att {attempt}): {str(e)}")
                if attempt == max_retries:
                    self.log_action(f"ERROR: Max retries append meta audit {file_path}")
                    return

    def sync_meta_audit_streams(self, sorted_entries):
        if not sorted_entries:
            return
        all_unique_entries = set()
        file_paths = [entry[4] for entry in sorted_entries]
        for file_path in file_paths:
            all_unique_entries.update(self.read_meta_audit(file_path))
        if not all_unique_entries:
            return
        entries_to_append_list = sorted(list(all_unique_entries))
        for file_path in file_paths:
            self.append_meta_audit(file_path, entries_to_append_list)

    def edit_meta_tag(self):
        selected = self.tree.currentItem()
        if not selected:
            self.log_action("WARN: Edit meta tag no selection.")
            return
        file_path = selected.data(0, Qt.ItemDataRole.UserRole + 1)  # Updated role
        current_meta = selected.text(5)
        if file_path is None:
            self.log_action("ERROR: No path for meta edit.")
            QMessageBox.warning(self, "Error", "No path.")
            return
        dialog = MetaTagEditor(file_path, current_meta, self.log_action, self)
        if dialog.exec():
            new_meta = self.read_ads_metadata(file_path)
            selected.setText(5, new_meta)
            self.log_action(f"INFO: Updated meta tag for {os.path.basename(file_path)}")

    def view_metadata_history(self, item):
        file_path = item.data(0, Qt.ItemDataRole.UserRole + 1)  # Updated role
        if file_path is None:
            self.log_action("ERROR: No path for history view.")
            QMessageBox.warning(self, "Error", "No path.")
            return
        dialog = MetadataHistoryDialog(file_path, self.log_action, self)
        dialog.exec()

    def select_main_icon(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Main Bar Icon", BASE_DIR, "Icon Files (*.png *.ico);;All Files (*.*)")
        if file_path:
            try:
                new_icon = QIcon(file_path)
                assert not new_icon.isNull()
                self.setWindowIcon(new_icon)
                self.save_config('main_icon_path', file_path)
                self.main_icon_path = file_path
                self.log_action(f"INFO: Set main icon: {file_path}")
            except:
                self.log_action(f"ERROR: Set main icon fail: {file_path}")
                QMessageBox.warning(self, "Error", "Invalid icon.")

    def select_tray_icon(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Tray Icon", BASE_DIR, "Icon Files (*.png *.ico);;All Files (*.*)")
        if file_path:
            try:
                new_icon = QIcon(file_path)
                assert not new_icon.isNull()
                self.tray_icon.setIcon(new_icon)
                self.save_config('tray_icon_path', file_path)
                self.tray_icon_path = file_path
                self.log_action(f"INFO: Set tray icon: {file_path}")
            except:
                self.log_action(f"ERROR: Set tray icon fail: {file_path}")
                QMessageBox.warning(self, "Error", "Invalid icon.")

    def update_tray_tooltip(self):
        selected_items = self.master_list.selectedItems()
        if selected_items:
            master_item_data = selected_items[0].data(1, Qt.ItemDataRole.UserRole)
            if master_item_data:
                self.current_document = master_item_data
                tooltip = f"{APP_NAME} - {self.current_document}"
            else:
                self.current_document = None
                tooltip = f"{APP_NAME} - Error"
                self.log_action("WARNING: No data in master item tooltip.")
        else:
            self.current_document = None
            tooltip = APP_NAME
        self.tray_icon.setToolTip(tooltip)

    def explore_production_directory(self):
        self.set_production_directory()
        self.tree.clear()
        self.preview_btn.setEnabled(False)

    def minimize_to_tray(self):
        self.hide()
        try:
            self.tray_icon.showMessage(APP_NAME, "Minimized. Right-click icon to restore.", QSystemTrayIcon.MessageIcon.Information, 2000)
        except Exception as e:
            self.log_action(f"WARNING: Tray message fail: {str(e)}")

    def restore_from_tray(self):
        self.show()
        self.activateWindow()

    def quit_application(self):
        self.log_action("INFO: App quitting.")
        self.tray_icon.hide()
        QApplication.quit()

    def closeEvent(self, event):
        self.quit_application()
        event.accept()

    def set_production_directory(self):
        current_dir = self.production_directory or BASE_DIR
        dir_path = QFileDialog.getExistingDirectory(self, "Select Production Directory", current_dir)
        if dir_path:
            self.production_directory = dir_path
            self.save_config('production_directory', dir_path)
            self.log_action(f"INFO: Prod dir set: {dir_path}")
            self.prod_dir_label.setText(dir_path)
            self.load_master_documents()
            self.tree.clear()
            self.preview_btn.setEnabled(False)

    def set_backup_directory(self):
        current_dir = self.backup_directory or BASE_DIR
        dir_path = QFileDialog.getExistingDirectory(self, "Select Backup Directory", current_dir)
        if dir_path:
            self.backup_directory = dir_path
            self.save_config('backup_directory', dir_path)
            self.log_action(f"INFO: Backup dir set: {dir_path}")
            self.backup_dir_label.setText(dir_path)
            self.load_master_documents()
            self.tree.clear()
            self.preview_btn.setEnabled(False)

    def show_context_menu(self, position: QPoint):
        item = self.tree.itemAt(position)
        if not item:
            return
        column = self.tree.columnAt(self.tree.viewport().mapFromGlobal(self.tree.mapToGlobal(position)).x())
        if column != 5:
            return
        menu = QMenu(self)
        view_history_action = QAction("View Metadata History", self)
        view_history_action.triggered.connect(lambda checked=False, it=item: self.view_metadata_history(it))
        menu.addAction(view_history_action)
        menu.exec(self.tree.mapToGlobal(position))

    def handle_tree_double_click(self, item, column):
        if column == 5:
            self.edit_meta_tag()
        else:
            self.use_selected()

    def refresh_meta_tags(self):
        self.log_action("INFO: Refreshing meta tags...")
        changed = False
        try:
            for i in range(self.tree.topLevelItemCount()):
                item = self.tree.topLevelItem(i)
                file_path = item.data(0, Qt.ItemDataRole.UserRole + 1)  # Updated role
                if file_path is None:
                    continue
                old_meta = item.text(5)
                new_meta = self.read_ads_metadata(file_path)
                if old_meta != new_meta:
                    item.setText(5, new_meta)
                    changed = True
            if changed:
                self.log_action("INFO: Meta Tags refreshed (changes).")
            else:
                self.log_action("INFO: Meta Tags refreshed (no changes).")
        except Exception as e:
            self.log_action(f"ERROR: Meta Tag refresh: {str(e)}")
            QMessageBox.warning(self, "Error", "Failed refresh.")

    def refresh_files(self):
        selected_master_items = self.master_list.selectedItems()
        if selected_master_items:
            self.log_action("INFO: Refreshing versions.")
            self.load_versions(selected_master_items[0])
        else:
            self.log_action("INFO: No master doc selected for refresh.")
            QMessageBox.information(self, "Refresh", "Select master file first.")

    def update_preview_button(self):
        selected = self.tree.currentItem()
        self.preview_btn.setEnabled(bool(selected))

    def preview_selected(self):
        selected = self.tree.currentItem()
        if not selected:
            self.log_action("INFO: Preview no selection.")
            QMessageBox.information(self, "No Selection", "Select version.")
            return
        path = selected.data(0, Qt.ItemDataRole.UserRole + 1)  # Updated role
        if path is None:
            self.log_action("ERROR: No path for preview.")
            QMessageBox.warning(self, "Error", "No file path.")
            return
        self.log_action(f"INFO: Previewing: {path}")
        try:
            dialog = PreviewDialog(path, self)
            dialog.exec()
        except Exception as e:
            self.log_action(f"ERROR: Preview dialog: {str(e)}")
            QMessageBox.warning(self, "Error", f"Preview failed:\n{str(e)}")

    def use_selected(self):
        selected = self.tree.currentItem()
        if not selected:
            self.log_action("INFO: Use no selection.")
            QMessageBox.information(self, "No Selection", "Select version.")
            return
        path = selected.data(0, Qt.ItemDataRole.UserRole + 1)  # Updated role
        version = selected.text(2)
        if path is None:
            self.log_action("ERROR: No path for use.")
            QMessageBox.warning(self, "Error", "No file path.")
            return
        if not os.path.exists(NOTEPADPP_PATH):
            self.log_action(f"ERROR: Npp not found: {NOTEPADPP_PATH}")
            QMessageBox.critical(self, "Npp Not Found", f"Npp not found:\n{NOTEPADPP_PATH}")
            return
        if not os.path.exists(path):
            self.log_action(f"ERROR: Backup file missing: {path}")
            QMessageBox.critical(self, "File Not Found", f"Backup file missing:\n{path}")
            return
        try:
            subprocess.Popen([NOTEPADPP_PATH, path])
            self.log_action(f"INFO: Opened Npp: {path}")
            QMessageBox.information(self, "File Opened", f"Opened {version}:\n{os.path.basename(path)}")
        except Exception as e:
            self.log_action(f"ERROR: Npp launch: {str(e)}")
            QMessageBox.critical(self, "Launch Error", f"Failed Npp launch:\n{str(e)}")

    def open_log_file(self):
        log_path = self.action_log_file
        if not os.path.exists(log_path):
            self.log_action(f"WARNING: Log missing: {log_path}.")
            QMessageBox.information(self, "Log File", f"Log will be created:\n{log_path}")
            return
        if not os.path.exists(NOTEPADPP_PATH):
            self.log_action(f"ERROR: Npp not found: {NOTEPADPP_PATH}")
            QMessageBox.critical(self, "Npp Not Found", f"Npp not found:\n{NOTEPADPP_PATH}")
            return
        try:
            subprocess.Popen([NOTEPADPP_PATH, log_path])
            self.log_action(f"INFO: Opened log: {log_path}")
        except Exception as e:
            self.log_action(f"ERROR: Open log: {str(e)}")
            QMessageBox.critical(self, "Launch Error", f"Failed open log:\n{str(e)}")

# --- Main Execution ---
if __name__ == "__main__":
    try:
        from ctypes import windll
        myappid = f'Friedberg.SnapVersionPlus.DocVExplorer.2'
        windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except:
        pass
    app = QApplication(sys.argv)
    viewer = DocumentVersionExplorer()
    viewer.show()
    sys.exit(app.exec())