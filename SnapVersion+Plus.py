import os
import sys
import json
import subprocess
import re
import ctypes
import difflib
from datetime import datetime
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QFileDialog, QTreeWidget, QTreeWidgetItem,
    QPushButton, QHBoxLayout, QVBoxLayout, QWidget, QMessageBox, QDialog,
    QTextEdit, QDialogButtonBox, QMenuBar, QSlider, QLabel, QDialog, QMenu
)
from PyQt6.QtGui import QAction, QIcon
from PyQt6.QtCore import Qt, QPoint
from DocumentVersionExplorer import DocumentVersionExplorer

# SnapVersion+Plus
# Determine the base directory dynamically (works for both script and packaged .exe)
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Constants for file paths and log file names
SCRIPT_NAME = os.path.splitext(os.path.basename(__file__))[0]
APP_NAME = "SnapVersion+Plus"
LOG_FILE = os.path.join(BASE_DIR, f"{SCRIPT_NAME}_log.json")
HELP_FILE = os.path.join(BASE_DIR, "help.txt")
NOTEPADPP_PATH = r"C:\Program Files\Notepad++\notepad++.exe"

class TransparencyDialog(QDialog):
    """Dialog to adjust the preview opacity with a slider."""
    def __init__(self, current_opacity, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Adjust Preview Transparency")
        self.setFixedSize(300, 150)

        layout = QVBoxLayout()

        # Slider for opacity (0% to 100%)
        self.slider = QSlider(Qt.Orientation.Horizontal)
        self.slider.setMinimum(0)
        self.slider.setMaximum(100)
        self.slider.setValue(int(current_opacity * 100))  # Convert 0.0-1.0 to 0-100
        self.slider.setTickPosition(QSlider.TickPosition.TicksBelow)
        self.slider.setTickInterval(10)
        layout.addWidget(QLabel("Transparency (0% = fully transparent, 100% = fully opaque):"))
        layout.addWidget(self.slider)

        # OK and Cancel buttons
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.setLayout(layout)

    def get_opacity(self):
        """Return the selected opacity value (0.0 to 1.0)."""
        return self.slider.value() / 100.0

class PreviewDialog(QDialog):
    """Dialog to preview the contents of a selected backup file."""
    def __init__(self, file_path, opacity, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Preview File")
        self.resize(600, 400)

        # Set the window opacity for transparency
        self.setWindowOpacity(opacity)

        layout = QVBoxLayout()
        
        # Read-only text area for file contents
        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                self.text_edit.setText(f.read())
        except Exception as e:
            self.text_edit.setText(f"Error reading file: {str(e)}")
        layout.addWidget(self.text_edit)

        # Close button
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.setLayout(layout)

class MetaTagEditor(QDialog):
    """Dialog for editing metadata in the NTFS ADS :source stream."""
    def __init__(self, file_path, current_meta, log_callback, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Meta Tag")
        self.resize(400, 150)
        self.file_path = file_path
        self.log_callback = log_callback

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
        """
        Save the edited metadata to the NTFS ADS :source stream, overwriting the existing content.
        Log the change to the :meta_audit stream (not the batch log).
        """
        try:
            # Get the new metadata content
            content = self.text_edit.toPlainText().strip()

            # Overwrite the :source stream with the latest metadata
            with open(f"{self.file_path}:source", "w", encoding="utf-8") as f:
                f.write(content)

            # Log metadata change only to dedicated stream :meta_audit (not the batch log)
            try:
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                history_entry = f"[{timestamp}] {content}\n"
                with open(f"{self.file_path}:meta_audit", "a", encoding="utf-8") as f:
                    f.write(history_entry)
            except Exception as e:
                self.log_callback(f"ERROR: Failed to write to metadata audit stream for {self.file_path}: {e}")

        except Exception as e:
            self.log_callback(f"ERROR: Failed to update metadata for {self.file_path}: {e}")
        self.accept()

class MetadataHistoryDialog(QDialog):
    """Dialog to display the metadata history for a file from its :meta_audit stream."""
    def __init__(self, file_path, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Metadata History")
        self.resize(400, 300)

        layout = QVBoxLayout()
        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)

        # Load metadata history from :meta_audit stream
        history = self.get_log_history_for_file(file_path)
        if history:
            self.text_edit.setPlainText("\n".join(history))
        else:
            self.text_edit.setPlainText("No metadata history available.")

        layout.addWidget(QLabel(f"Metadata history for:\n{os.path.basename(file_path)}"))
        layout.addWidget(self.text_edit)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.setLayout(layout)

    def get_log_history_for_file(self, file_path):
        """
        Load metadata history from the file's :meta_audit stream, sorted newest-first.

        Args:
            file_path (str): The path to the file.

        Returns:
            list: List of history entries, newest-first.
        """
        try:
            ads_path = f"{file_path}:meta_audit"
            with open(ads_path, "r", encoding="utf-8") as f:
                lines = f.readlines()
            # Sort newest-first (lines are appended, so reverse them)
            return [line.strip() for line in lines if line.strip()][::-1]
        except Exception:
            return []

class BackupViewer(QMainWindow):
    def __init__(self):
        """Initialize the BackupViewer application window and UI components."""
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(800, 450)

        # Ensure the base directory exists for logs
        try:
            os.makedirs(BASE_DIR, exist_ok=True)
        except Exception as e:
            print(f"ERROR: Failed to create base directory {BASE_DIR}: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to create base directory: {str(e)}\nApplication may not function correctly.")

        # Log application start (will use a default log until batch is selected)
        self.batch_log_file = os.path.join(BASE_DIR, f"{SCRIPT_NAME}_default_log.txt")
        self.log_action("INFO: Application started")

        # Default opacity for preview dialog (1.0 = fully opaque)
        self.preview_opacity = 1.0

        # Set the titlebar color (blue) with error handling
        try:
            self.set_titlebar_color()
        except Exception as e:
            error_msg = f"ERROR: Failed to set titlebar color: {str(e)}"
            self.log_action(error_msg)
            QMessageBox.warning(self, "Error", "Failed to set titlebar color. Continuing without custom titlebar.")

        # Load the JSON log file to retrieve the last selected directory and icon
        log_data = self.load_log()
        self.home_directory = log_data.get('directory', '')
        self.batch_file_name = None

        # Set the saved icon if it exists
        saved_icon_path = log_data.get('icon_path')
        if saved_icon_path and os.path.exists(saved_icon_path):
            try:
                icon = QIcon(saved_icon_path)
                if not icon.isNull():
                    self.setWindowIcon(icon)
                    self.log_action(f"INFO: Loaded saved icon from: {saved_icon_path}")
            except Exception as e:
                self.log_action(f"WARNING: Failed to load saved icon: {str(e)}")

        # Create the help file if it doesn't exist
        self.create_help_file()

        # Create the menu bar
        self.create_menu_bar()

        # Set up the main UI layout
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout()
        self.central_widget.setLayout(self.main_layout)

        # Add Preview button above the tree widget with blue styling
        self.preview_btn = QPushButton("Preview")
        self.preview_btn.setStyleSheet("background-color: #0078D7; color: white; font-weight: bold;")
        self.preview_btn.clicked.connect(self.preview_selected)
        self.preview_btn.setEnabled(False)
        self.main_layout.addWidget(self.preview_btn)

        # Create the tree widget for displaying backup files
        self.tree = QTreeWidget()
        self.tree.setColumnCount(6)  # Added Total Lines column
        self.tree.setHeaderLabels(["Date/Time", "Group Alias", "Version", "Changes", "Total Lines", "Meta Tag"])
        self.tree.setStyleSheet("QTreeWidget { padding: 5px; } QTreeWidget::item { padding: 5px; }")
        self.tree.setAlternatingRowColors(True)
        self.tree.itemDoubleClicked.connect(self.handle_tree_double_click)
        self.tree.itemSelectionChanged.connect(self.update_preview_button)
        # Add context menu for right-click on Meta Tag column
        self.tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self.show_context_menu)
        self.main_layout.addWidget(self.tree)

        # Set up the button layout with two rows
        self.button_container = QVBoxLayout()

        # First row: Set Directory, Select File Batch, Select File, Refresh Meta Tags, Refresh Files
        self.button_layout_top = QHBoxLayout()
        self.dir_btn = QPushButton("Set Directory")
        self.dir_btn.clicked.connect(self.set_directory)
        self.button_layout_top.addWidget(self.dir_btn)

        self.batch_btn = QPushButton("Select File Batch To Choose From")
        self.batch_btn.clicked.connect(self.set_batch_file)
        self.button_layout_top.addWidget(self.batch_btn)

        self.use_btn = QPushButton("Select File")
        self.use_btn.clicked.connect(self.use_selected)
        self.button_layout_top.addWidget(self.use_btn)

        self.refresh_meta_btn = QPushButton("Refresh Meta Tags")
        self.refresh_meta_btn.setStyleSheet("background-color: #0078D7; color: white; font-weight: bold;")
        self.refresh_meta_btn.clicked.connect(self.refresh_meta_tags)
        self.button_layout_top.addWidget(self.refresh_meta_btn)

        self.refresh_files_btn = QPushButton("Refresh Files")
        self.refresh_files_btn.setStyleSheet("background-color: #0078D7; color: white; font-weight: bold;")
        self.refresh_files_btn.clicked.connect(self.refresh_files)
        self.button_layout_top.addWidget(self.refresh_files_btn)

        # Second row: Close button
        self.button_layout_bottom = QHBoxLayout()
        self.close_btn = QPushButton("Close")
        self.close_btn.setStyleSheet("background-color: #a83232; color: white; font-weight: bold;")
        self.close_btn.clicked.connect(self.close)
        self.button_layout_bottom.addWidget(self.close_btn)

        # Add the two rows to the container
        self.button_container.addLayout(self.button_layout_top)
        self.button_container.addLayout(self.button_layout_bottom)
        self.main_layout.addLayout(self.button_container)

    def create_help_file(self):
        """Create a help file explaining the app's functionality and file locations."""
        help_content = """SnapVersion+Plus Help File
=================================

Purpose:
--------
SnapVersion+Plus is a utility to help you manage and restore Notepad++ backup files (.bak) in a user-friendly way. It allows you to select a directory containing backup files, view all backups for a specific file, preview their contents, and restore them in Notepad++.

How It Works:
-------------
1. **Set Directory**: Choose a directory containing your Notepad++ backup files (.bak).
2. **Select Batch**: Pick a file (any file or .bak) to list all related backup files in the selected directory.
3. **View Backups**: The app lists all .bak files with the same base name, sorted by creation date (newest first), showing:
   - Date/Time (e.g., "sat 05/03/2025 11:53am", reflecting the creation date of the backup)
   - Group Alias (base name of the file)
   - Version (V1 for the oldest, incrementing to the newest at the top, e.g., V9 for the newest if there are 9 versions)
   - Changes (number of lines changed compared to the previous older version)
   - Total Lines (total number of lines in the file)
   - Meta Tag (user-editable metadata stored in NTFS ADS :source stream)
4. **Preview**: Use the Preview button to view a file's contents in a popup without opening it in Notepad++.
5. **Edit Meta Tag**: Double-click the Meta Tag column to edit metadata for a backup file.
6. **View Metadata History**: Right-click the Meta Tag column and select "View Metadata History" to see all previous metadata changes for that file.
7. **Refresh Meta Tags**: Click the Refresh Meta Tags button to reload metadata for all displayed files.
8. **Refresh Files**: Click the Refresh Files button to reload the list of backup files, detecting new backups.
9. **Restore**: Double-click any other column or click "Select File" to open a backup in Notepad++. A confirmation popup will appear, and the app will remain open.
10. **Transparency**: Adjust the preview popup's transparency via Settings > Transparency (0% to 100%).
11. **Clear Log**: Clear the application log file via Settings > Clear Log.

File Locations and Purposes:
----------------------------
- **Log File (<app_name>_log.json)**:
  Location: Same directory as the app executable (e.g., C:\\path\\to\\exe\\<app_name>_log.json)
  Purpose: Stores persistent settings, such as the last selected directory.
- **Batch Log File (<batch_base_name>_loghistory.txt)**:
  Location: Same directory as the app executable (e.g., C:\\path\\to\\exe\\<batch_base_name>_loghistory.txt)
  Purpose: Logs all actions (e.g., file selections, restorations) specific to the selected batch.
- **Metadata Audit Stream (<file_path>:meta_audit)**:
  Location: NTFS ADS stream attached to each backup file
  Purpose: Stores the history of metadata changes for that specific file. Preserved across renames by copying entries from all versions to every file with 100% integrity.
- **Help File (help.txt)**:
  Location: Same directory as the app executable (e.g., C:\\path\\to\\exe\\help.txt)
  Purpose: This file, providing documentation on the app's functionality.
- **Notepad++ Executable**:
  Location: C:\\Program Files\\Notepad++\\notepad++.exe
  Purpose: Used to open selected backup files for restoration.

Notes:
------
- The app creates these files in the same directory as the executable if they don't exist.
- All errors are logged to the batch-specific log file for troubleshooting.
- The Meta Tag and metadata history features require an NTFS filesystem to store data in the :source and :meta_audit streams.
- Metadata history preservation includes robust fallback logic to ensure no entries are lost during renames, restores, or versioning.
"""
        try:
            os.makedirs(os.path.dirname(HELP_FILE), exist_ok=True)
            with open(HELP_FILE, 'w') as f:
                f.write(help_content)
            self.log_action(f"INFO: Created help file at {HELP_FILE}")
        except Exception as e:
            error_msg = f"ERROR: Failed to create help file {HELP_FILE}: {str(e)}"
            self.log_action(error_msg)
            QMessageBox.warning(self, "Error", "Failed to create help file. Help menu may not function correctly.")

    def create_menu_bar(self):
        """Create a Windows-style menu bar with File, Settings, and About options."""
        menu_bar = QMenuBar(self)
        self.setMenuBar(menu_bar)

        # File menu
        file_menu = menu_bar.addMenu("File")

        # Set Directory
        set_dir_action = QAction("Set Directory", self)
        set_dir_action.triggered.connect(self.set_directory)
        file_menu.addAction(set_dir_action)

        # Select Batch
        select_batch_action = QAction("Select Batch", self)
        select_batch_action.triggered.connect(self.set_batch_file)
        file_menu.addAction(select_batch_action)

        # Change Icon
        change_icon_action = QAction("Change Icon", self)
        change_icon_action.triggered.connect(self.change_icon)
        file_menu.addAction(change_icon_action)

        # Open Log
        open_log_action = QAction("Open Log", self)
        open_log_action.triggered.connect(self.open_log)
        file_menu.addAction(open_log_action)

        file_menu.addSeparator()

        # Close Window
        close_action = QAction("Close Window", self)
        close_action.triggered.connect(self.close)
        file_menu.addAction(close_action)

        # Close Full Application
        close_full_action = QAction("Close Full Application", self)
        close_full_action.triggered.connect(self.close_full_application)
        file_menu.addAction(close_full_action)

        # Settings menu
        settings_menu = menu_bar.addMenu("Settings")

        # Transparency option (with slider)
        transparency_action = QAction("Transparency", self)
        transparency_action.triggered.connect(self.adjust_transparency)
        settings_menu.addAction(transparency_action)

        # Clear Log option
        clear_log_action = QAction("Clear Log", self)
        clear_log_action.triggered.connect(self.clear_log)
        settings_menu.addAction(clear_log_action)

        # Tools menu
        tools_menu = menu_bar.addMenu("Tools")

        # Document Version Explorer
        doc_explorer_action = QAction("Document Version Explorer", self)
        doc_explorer_action.triggered.connect(self.open_document_explorer)
        tools_menu.addAction(doc_explorer_action)

        # About menu
        about_menu = menu_bar.addMenu("About")

        # About SnapVersion+Plus
        about_action = QAction(f"About {APP_NAME}", self)
        about_action.triggered.connect(self.show_about)
        about_menu.addAction(about_action)

        # Help (open help.txt)
        help_action = QAction("Help", self)
        help_action.triggered.connect(self.open_help)
        about_menu.addAction(help_action)

    def change_icon(self):
        """Open a file dialog to select a new application icon."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Icon File",
            BASE_DIR,
            "Icon Files (*.ico *.png);;All Files (*.*)"
        )
        if file_path:
            try:
                icon = QIcon(file_path)
                if not icon.isNull():
                    self.setWindowIcon(icon)
                    self.save_log('icon_path', file_path)
                    self.log_action(f"INFO: Changed application icon to: {file_path}")
                    QMessageBox.information(self, "Icon Changed", "Application icon has been updated.")
                else:
                    raise ValueError("Invalid icon file")
            except Exception as e:
                error_msg = f"ERROR: Failed to set icon: {str(e)}"
                self.log_action(error_msg)
                QMessageBox.warning(self, "Error", f"Failed to set icon: {str(e)}")

    def show_context_menu(self, position: QPoint):
        """
        Show a context menu when right-clicking on the tree widget.

        Args:
            position (QPoint): The position of the right-click.
        """
        # Get the item and column at the clicked position
        item = self.tree.itemAt(position)
        if not item:
            return

        column = self.tree.columnAt(self.tree.viewport().mapFromGlobal(self.tree.mapToGlobal(position)).x())
        if column != 5:  # Only show menu on Meta Tag column (column 5)
            return

        # Create context menu
        menu = QMenu(self)
        view_history_action = QAction("View Metadata History", self)
        view_history_action.triggered.connect(lambda: self.view_metadata_history(item))
        menu.addAction(view_history_action)

        # Show the menu at the cursor position
        menu.exec(self.tree.mapToGlobal(position))

    def view_metadata_history(self, item):
        """
        Open a dialog to view the metadata history for the selected file.

        Args:
            item (QTreeWidgetItem): The selected item in the tree widget.
        """
        file_path = item.data(0, Qt.ItemDataRole.UserRole)
        dialog = MetadataHistoryDialog(file_path, self)
        dialog.exec()

    def adjust_transparency(self):
        """Open a dialog to adjust the preview opacity with a slider."""
        dialog = TransparencyDialog(self.preview_opacity, self)
        if dialog.exec():
            new_opacity = dialog.get_opacity()
            self.preview_opacity = new_opacity
            self.log_action(f"INFO: Set preview opacity to {new_opacity}")

    def open_help(self):
        """Open the help file in Notepad++."""
        if not os.path.exists(HELP_FILE):
            self.log_action(f"ERROR: Help file {HELP_FILE} does not exist")
            QMessageBox.warning(self, "Error", "Help file not found.")
            return
        if not os.path.exists(NOTEPADPP_PATH):
            error_msg = f"ERROR: Notepad++ not found at {NOTEPADPP_PATH}"
            self.log_action(error_msg)
            QMessageBox.critical(self, "Error", f"Notepad++ not found at {NOTEPADPP_PATH}. Please ensure it is installed.")
            return
        try:
            subprocess.Popen([NOTEPADPP_PATH, HELP_FILE])
            self.log_action(f"INFO: Opened help file: {HELP_FILE}")
        except (FileNotFoundError, OSError) as e:
            error_msg = f"ERROR: Failed to open help file {HELP_FILE}: {str(e)}"
            self.log_action(error_msg)
            QMessageBox.critical(self, "Error", f"Failed to open help file: {str(e)}")

    def show_about(self):
        """Show the About popup with app details."""
        about_text = (
            f"{APP_NAME}\n"
            "Author: Christopher Friedberg\n"
            "Date: May 03, 2025\n"
            "Version: 1.0.0\n\n"
            "Usage is open but limited in scope to personal and educational purposes. "
            "Not intended for commercial use without prior consent from the author."
        )
        QMessageBox.information(self, f"About {APP_NAME}", about_text)

    def open_log(self):
        """Open the batch-specific application log file in Notepad++ for viewing."""
        if not os.path.exists(self.batch_log_file):
            self.log_action(f"ERROR: Log file {self.batch_log_file} does not exist")
            QMessageBox.warning(self, "Error", "Log file not found. Select a batch file to generate a log.")
            return
        if not os.path.exists(NOTEPADPP_PATH):
            error_msg = f"ERROR: Notepad++ not found at {NOTEPADPP_PATH}"
            self.log_action(error_msg)
            QMessageBox.critical(self, "Error", f"Notepad++ not found at {NOTEPADPP_PATH}. Please ensure it is installed.")
            return
        try:
            subprocess.Popen([NOTEPADPP_PATH, self.batch_log_file])
            self.log_action(f"INFO: Opened log file: {self.batch_log_file}")
        except (FileNotFoundError, OSError) as e:
            error_msg = f"ERROR: Failed to open log file {self.batch_log_file}: {str(e)}"
            self.log_action(error_msg)
            QMessageBox.critical(self, "Error", f"Failed to open log file: {str(e)}")

    def clear_log(self):
        """Clear the batch-specific application log file (truncate to empty)."""
        try:
            with open(self.batch_log_file, 'w') as f:
                f.write("")
            self.log_action("INFO: Cleared application log file")
            QMessageBox.information(self, "Log Cleared", "Application log has been cleared.")
        except Exception as e:
            error_msg = f"ERROR: Failed to clear log file {self.batch_log_file}: {str(e)}"
            self.log_action(error_msg)
            QMessageBox.warning(self, "Error", "Failed to clear log file.")

    def log_action(self, message):
        """
        Log an action or error to the batch-specific log file with a timestamp.

        Args:
            message (str): The message to log (e.g., "INFO: Directory set" or "ERROR: Failed to...").
        """
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        try:
            os.makedirs(os.path.dirname(self.batch_log_file), exist_ok=True)
            with open(self.batch_log_file, 'a') as f:
                f.write(log_entry)
                f.flush()
                os.fsync(f.fileno())
        except Exception as e:
            print(f"ERROR: Failed to write to log file {self.batch_log_file}: {str(e)}")

    def set_titlebar_color(self):
        """Set the window titlebar color to blue using Windows API."""
        hwnd = int(self.winId())
        DWMWA_USE_IMMERSIVE_DARK_MODE = 20
        DWMWA_ACCENT_POLICY = 19

        class ACCENT_POLICY(ctypes.Structure):
            _fields_ = [
                ("AccentState", ctypes.c_uint),
                ("AccentFlags", ctypes.c_uint),
                ("GradientColor", ctypes.c_uint),
                ("AnimationId", ctypes.c_uint),
            ]

        accent = ACCENT_POLICY()
        accent.AccentState = 3
        accent.GradientColor = 0xFF0000

        dwmapi = ctypes.WinDLL("dwmapi")
        dwmapi.DwmSetWindowAttribute(hwnd, DWMWA_ACCENT_POLICY, ctypes.byref(accent), ctypes.sizeof(accent))

        dark_mode = ctypes.c_int(1)
        dwmapi.DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ctypes.byref(dark_mode), ctypes.sizeof(dark_mode))

    def read_ads_metadata(self, file_path):
        """
        Reads metadata from NTFS Alternate Data Stream :source.

        Args:
            file_path (str): The path to the file.

        Returns:
            str: The metadata content, or an empty string if not found or on error.
        """
        try:
            ads_path = f"{file_path}:source"
            with open(ads_path, "r", encoding="utf-8") as f:
                return f.read().strip()
        except Exception:
            return ""

    def read_meta_audit(self, file_path):
        """
        Reads the :meta_audit stream for a file with retry logic.

        Args:
            file_path (str): The path to the file.

        Returns:
            list: List of audit entries, or empty list if not found or on error.
        """
        max_retries = 3
        retry_count = 0
        while retry_count < max_retries:
            try:
                ads_path = f"{file_path}:meta_audit"
                with open(ads_path, "r", encoding="utf-8") as f:
                    lines = f.readlines()
                self.log_action(f"INFO: Successfully read :meta_audit stream for {file_path}")
                return [line.strip() for line in lines if line.strip()]
            except Exception as e:
                retry_count += 1
                self.log_action(f"ERROR: Failed to read :meta_audit stream for {file_path} (attempt {retry_count}/{max_retries}): {str(e)}")
                if retry_count == max_retries:
                    # Fallback: Attempt to copy to a temporary file and read from there
                    try:
                        temp_path = os.path.join(BASE_DIR, f"temp_meta_audit_{os.path.basename(file_path)}.txt")
                        with open(ads_path, "rb") as f_in, open(temp_path, "wb") as f_out:
                            f_out.write(f_in.read())
                        with open(temp_path, "r", encoding="utf-8") as f:
                            lines = f.readlines()
                        os.remove(temp_path)
                        self.log_action(f"INFO: Successfully read :meta_audit stream via temp file for {file_path}")
                        return [line.strip() for line in lines if line.strip()]
                    except Exception as e2:
                        self.log_action(f"ERROR: Fallback failed to read :meta_audit stream for {file_path}: {str(e2)}")
                        return []
                continue
        return []

    def append_meta_audit(self, file_path, entries):
        """
        Appends audit entries to the :meta_audit stream of a file with retry logic and duplicate checking.

        Args:
            file_path (str): The path to the file.
            entries (list): List of audit entries to append.
        """
        if not entries:
            return

        # Read existing entries to avoid duplicates
        existing_entries = set(self.read_meta_audit(file_path))
        new_entries = [entry for entry in entries if entry.strip() and entry.strip() not in existing_entries]

        if not new_entries:
            self.log_action(f"INFO: No new :meta_audit entries to append for {file_path}")
            return

        max_retries = 3
        retry_count = 0
        while retry_count < max_retries:
            try:
                ads_path = f"{file_path}:meta_audit"
                with open(ads_path, "a", encoding="utf-8") as f:
                    f.writelines(f"{entry}\n" for entry in new_entries)
                self.log_action(f"INFO: Successfully appended {len(new_entries)} new :meta_audit entries to {file_path}")
                return
            except Exception as e:
                retry_count += 1
                self.log_action(f"ERROR: Failed to append to :meta_audit for {file_path} (attempt {retry_count}/{max_retries}): {str(e)}")
                if retry_count == max_retries:
                    # Fallback: Write to a temporary file and attempt to merge
                    try:
                        temp_path = os.path.join(BASE_DIR, f"temp_meta_audit_{os.path.basename(file_path)}.txt")
                        with open(temp_path, "a", encoding="utf-8") as f:
                            f.writelines(f"{entry}\n" for entry in new_entries)
                        with open(ads_path, "ab") as f_out:
                            with open(temp_path, "rb") as f_in:
                                f_out.write(f_in.read())
                        os.remove(temp_path)
                        self.log_action(f"INFO: Successfully appended :meta_audit entries via temp file for {file_path}")
                    except Exception as e2:
                        self.log_action(f"ERROR: Fallback failed to append :meta_audit for {file_path}: {str(e2)}")
                    return
                continue

    def edit_meta_tag(self):
        """Open a dialog to edit the metadata for the selected backup file."""
        selected = self.tree.currentItem()
        if not selected:
            return
        file_path = selected.data(0, Qt.ItemDataRole.UserRole)
        current_meta = selected.text(5)  # Meta Tag column (column 5)
        dialog = MetaTagEditor(file_path, current_meta, self.log_action, self)
        if dialog.exec():
            new_meta = self.read_ads_metadata(file_path)
            selected.setText(5, new_meta)

    def handle_tree_double_click(self, item, column):
        """
        Handle double-click events on the tree widget.
        Edits metadata if the Meta Tag column is clicked; otherwise, opens the file in Notepad++.

        Args:
            item (QTreeWidgetItem): The clicked item.
            column (int): The clicked column (0-based index).
        """
        if column == 5:  # Meta Tag column (column 5)
            self.edit_meta_tag()
        else:
            self.use_selected()

    def refresh_meta_tags(self):
        """
        Refresh the Meta Tag column for all displayed backup files by re-reading the :source stream.
        """
        try:
            for i in range(self.tree.topLevelItemCount()):
                item = self.tree.topLevelItem(i)
                file_path = item.data(0, Qt.ItemDataRole.UserRole)
                new_meta = self.read_ads_metadata(file_path)
                item.setText(5, new_meta)  # Meta Tag column (column 5)
            self.log_action("INFO: Refreshed Meta Tags for all displayed files")
        except Exception as e:
            error_msg = f"ERROR: Failed to refresh Meta Tags: {str(e)}"
            self.log_action(error_msg)
            QMessageBox.warning(self, "Error", "Failed to refresh Meta Tags.")

    def refresh_files(self):
        """
        Refresh the list of backup files by re-running the load logic for the current batch file.
        """
        if not self.batch_file_name:
            self.log_action("INFO: No batch file selected to refresh")
            QMessageBox.information(self, "No Batch File", "Please select a batch file to refresh.")
            return
        try:
            self.load_backups_from_name()
            self.log_action(f"INFO: Refreshed file list for batch: {self.batch_file_name}")
        except Exception as e:
            error_msg = f"ERROR: Failed to refresh file list: {str(e)}"
            self.log_action(error_msg)
            QMessageBox.warning(self, "Error", "Failed to refresh file list.")

    def load_log(self):
        """
        Load the JSON log file for directory persistence.

        Returns:
            dict: The loaded log data, or an empty dict if loading fails.
        """
        if not os.path.exists(LOG_FILE):
            self.log_action(f"INFO: Log file {LOG_FILE} does not exist, starting fresh")
            return {}
        try:
            with open(LOG_FILE, 'r') as f:
                return json.load(f)
        except json.JSONDecodeError as e:
            error_msg = f"ERROR: Corrupted log file {LOG_FILE}: {str(e)}"
            self.log_action(error_msg)
            QMessageBox.warning(self, "Error", "Log file is corrupted. Starting with default settings.")
            return {}
        except Exception as e:
            error_msg = f"ERROR: Failed to load log file {LOG_FILE}: {str(e)}"
            self.log_action(error_msg)
            QMessageBox.warning(self, "Error", "Failed to load log file. Starting with default settings.")
            return {}

    def save_log(self, key, value):
        """
        Save a key-value pair to the JSON log file.

        Args:
            key (str): The key to save (e.g., 'directory').
            value (str): The value to save (e.g., the directory path).
        """
        data = self.load_log()
        data[key] = value
        try:
            os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
            with open(LOG_FILE, 'w') as f:
                json.dump(data, f)
            self.log_action(f"INFO: Saved {key} to log file: {value}")
        except Exception as e:
            error_msg = f"ERROR: Failed to save log file {LOG_FILE}: {str(e)}"
            self.log_action(error_msg)
            QMessageBox.warning(self, "Error", "Failed to save directory settings. Changes may not persist.")

    def set_directory(self):
        """Set the directory for searching backup files via a file dialog."""
        try:
            dir_path = QFileDialog.getExistingDirectory(self, "Select Directory")
            if dir_path:
                self.home_directory = dir_path
                self.save_log('directory', dir_path)
                self.tree.clear()
                self.batch_file_name = None
                self.batch_log_file = os.path.join(BASE_DIR, f"{SCRIPT_NAME}_default_log.txt")
                self.log_action(f"INFO: Directory set to {dir_path}")
                QMessageBox.information(self, "Directory Set", f"Root directory saved:\n{dir_path}")
        except Exception as e:
            error_msg = f"ERROR: Failed to set directory: {str(e)}"
            self.log_action(error_msg)
            QMessageBox.warning(self, "Error", "Failed to set directory. Please try again.")

    def set_batch_file(self):
        """
        Select a file to determine the batch of backups to display.
        Opens a file dialog in the user-selected directory.
        Updates the batch-specific log file path.
        """
        if not self.home_directory or not os.path.exists(self.home_directory):
            self.log_action("ERROR: No valid directory set for batch file selection")
            QMessageBox.warning(self, "No Directory", "Please set a valid root directory first.")
            return
        try:
            file_path, _ = QFileDialog.getOpenFileName(
                self, 
                "Select Batch File", 
                self.home_directory,
                "All Files (*.*);;Notepad++ Backups (*.bak)"
            )
            if file_path:
                self.batch_file_name = file_path
                self.save_log('batch_file', file_path)
                # Update the batch-specific log file based on the base name
                base_name = self._get_base_name(os.path.basename(file_path))
                self.batch_log_file = os.path.join(BASE_DIR, f"{base_name}_loghistory.txt")
                self.log_action(f"INFO: Batch file selected: {file_path}")
                self.load_backups_from_name()
        except Exception as e:
            error_msg = f"ERROR: Failed to select batch file: {str(e)}"
            self.log_action(error_msg)
            QMessageBox.warning(self, "Error", "Failed to select batch file. Please try again.")

    def _get_base_name(self, filename):
        """
        Extract the base name of a file by removing all extensions for non-.bak files or the timestamp for .bak files.

        Args:
            filename (str): The name of the file.

        Returns:
            str: The base name of the file.
        """
        # First, handle .bak files with timestamp
        pattern = r'^(.*)\.\d{4}-\d{2}-\d{2}_\d{6}\.bak$'
        match = re.match(pattern, filename)
        if match:
            base = match.group(1)
        else:
            # For non-.bak files, remove all extensions by splitting on dots
            base = filename.split('.')[0]
        # If the base still contains extensions (e.g., script.py.py), strip them all
        while '.' in base:
            base = base.split('.')[0]
        return base

    def load_backups_from_name(self):
        """
        Load and display backup files from the selected directory.
        Matches files with the same base name as the selected file using improved regex.
        Compares each file with the next (older) version to count differences.
        Uses creation time for sorting (newest first) and labels oldest as V1, newest as highest version.
        Preserves :meta_audit stream across all files with robust retry and logging.
        """
        if not self.batch_file_name:
            self.log_action("ERROR: No batch file selected")
            return

        if not os.path.exists(self.home_directory):
            error_msg = f"ERROR: Selected directory not found: {self.home_directory}"
            self.log_action(error_msg)
            QMessageBox.warning(self, "Error", f"Selected directory not found: {self.home_directory}")
            return

        filename = os.path.basename(self.batch_file_name)
        # Extract base name (e.g., 'script' from 'script.py' or 'script.py.2025-05-03_115301.bak')
        base = self._get_base_name(filename)

        try:
            # Find all .bak files in the directory that start with the base name
            backups = [f for f in os.listdir(self.home_directory) if f.startswith(base) and f.endswith('.bak')]
        except Exception as e:
            error_msg = f"ERROR: Failed to list files in directory {self.home_directory}: {str(e)}"
            self.log_action(error_msg)
            QMessageBox.warning(self, "Error", f"Failed to access directory: {self.home_directory}")
            return

        self.tree.clear()

        if not backups:
            self.log_action(f"INFO: No backups found for: {base} in {self.home_directory}")
            QMessageBox.information(self, "No Backups", f"No backups found for: {base} in {self.home_directory}")
            return

        entries = []
        for name in backups:
            path = os.path.join(self.home_directory, name)
            try:
                # Use creation time for sorting
                timestamp = datetime.fromtimestamp(os.path.getctime(path))
                dt_str = timestamp.strftime("%a %m/%d/%Y %I:%M%p").lower()
                entries.append((timestamp, dt_str, base, name, path))
            except OSError as e:
                error_msg = f"ERROR: Failed to access {path}: {str(e)}"
                self.log_action(error_msg)
                continue

        # Sort by creation timestamp (newest first)
        entries.sort(key=lambda x: x[0], reverse=True)

        # Collect all :meta_audit entries from all files in the batch
        all_meta_audit_entries = []
        for entry in entries:
            file_path = entry[4]
            meta_audit_entries = self.read_meta_audit(file_path)
            if meta_audit_entries:
                all_meta_audit_entries.extend(meta_audit_entries)
            else:
                self.log_action(f"INFO: No :meta_audit entries found for {file_path}")

        # Append all collected :meta_audit entries to every file in the batch
        for entry in entries:
            file_path = entry[4]
            self.append_meta_audit(file_path, all_meta_audit_entries)

        diff_counts = []
        total_lines = []
        for i in range(len(entries)):
            # Count total lines in the current file
            try:
                with open(entries[i][4], 'r', encoding='utf-8') as f:
                    line_count = len(f.readlines())
                total_lines.append(str(line_count))
            except Exception as e:
                error_msg = f"ERROR: Failed to count lines in {entries[i][4]}: {str(e)}"
                self.log_action(error_msg)
                total_lines.append("Error")

            # Calculate changes by comparing with the next (older) version
            if i == len(entries) - 1:  # Oldest file (V1)
                diff_counts.append("N/A")  # No older version to compare with
            else:
                curr_path = entries[i][4]    # Current file
                next_path = entries[i+1][4]  # Next (older) file
                try:
                    with open(curr_path, 'r', encoding='utf-8') as f_curr, open(next_path, 'r', encoding='utf-8') as f_next:
                        curr_lines = f_curr.readlines()
                        next_lines = f_next.readlines()
                        diff = list(difflib.unified_diff(next_lines, curr_lines, lineterm=''))  # Compare older to newer
                        change_count = sum(1 for line in diff if line.startswith('+') or line.startswith('-'))
                        # Adjust for unified diff format: exclude header lines (---, +++, @@)
                        change_count = max(0, change_count - 3)
                        # Determine if it's an addition or subtraction based on total lines
                        curr_total = len(curr_lines)
                        next_total = len(next_lines)
                        if curr_total > next_total:
                            change_sign = "+"
                        elif curr_total < next_total:
                            change_sign = "-"
                        else:
                            change_sign = ""
                        diff_counts.append(f"{change_sign}{change_count} lines")
                except Exception as e:
                    error_msg = f"ERROR: Failed to compare {curr_path} with {next_path}: {str(e)}"
                    self.log_action(error_msg)
                    diff_counts.append("Error")

        # Version numbers: V1 for oldest (bottom), incrementing to newest (top)
        total_versions = len(entries)
        for idx, (entry, diff_count, total_line_count) in enumerate(zip(entries, diff_counts, total_lines)):
            version_number = total_versions - idx  # Newest (top) gets highest number, oldest (bottom) gets V1
            _, dt, base, name, full = entry
            meta = self.read_ads_metadata(full)
            item = QTreeWidgetItem([dt, base, f"V{version_number}", diff_count, total_line_count, meta])
            item.setData(0, Qt.ItemDataRole.UserRole, full)
            self.tree.addTopLevelItem(item)

        try:
            for i in range(self.tree.columnCount()):
                self.tree.resizeColumnToContents(i)
        except Exception as e:
            error_msg = f"ERROR: Failed to resize tree columns: {str(e)}"
            self.log_action(error_msg)
            QMessageBox.warning(self, "Error", "Failed to adjust display. List may not display correctly.")

    def update_preview_button(self):
        """Enable or disable the Preview button based on whether a file is selected."""
        selected = self.tree.currentItem()
        self.preview_btn.setEnabled(bool(selected))

    def preview_selected(self):
        """Open a preview dialog for the selected backup file with the current opacity setting."""
        selected = self.tree.currentItem()
        if not selected:
            self.log_action("INFO: No version selected for preview")
            QMessageBox.information(self, "No Selection", "Select a version to preview.")
            return
        path = selected.data(0, Qt.ItemDataRole.UserRole)
        try:
            dialog = PreviewDialog(path, self.preview_opacity, self)
            dialog.exec()
        except Exception as e:
            error_msg = f"ERROR: Failed to preview file {path}: {str(e)}"
            self.log_action(error_msg)
            QMessageBox.warning(self, "Error", f"Failed to preview file: {str(e)}")

    def use_selected(self):
        """
        Open the selected backup file in Notepad++ when double-clicked or 'Select File' is clicked.
        Show a confirmation popup but do not close the app.
        """
        selected = self.tree.currentItem()
        if not selected:
            self.log_action("INFO: No version selected to open")
            QMessageBox.information(self, "No Selection", "Select a version to open.")
            return
        path = selected.data(0, Qt.ItemDataRole.UserRole)
        version = selected.text(2)
        self.open_with_notepadpp(path, version)

    def open_with_notepadpp(self, path, version):
        """
        Open the specified file in Notepad++ using the standard install path.
        Show a confirmation popup after opening.

        Args:
            path (str): The file path to open.
            version (str): The version label (e.g., "V1").
        """
        if not os.path.exists(NOTEPADPP_PATH):
            error_msg = f"ERROR: Notepad++ not found at {NOTEPADPP_PATH}"
            self.log_action(error_msg)
            QMessageBox.critical(self, "Error", f"Notepad++ not found at {NOTEPADPP_PATH}. Please ensure it is installed.")
            return

        try:
            subprocess.Popen([NOTEPADPP_PATH, path])
            self.log_action(f"INFO: Selected file: {path}")
            # Preserve :meta_audit stream for the restored file if necessary
            # Since this is a restore, the file might be saved under a new name; we'll handle this in the refresh
            QMessageBox.information(
                self,
                "Restore Complete",
                f"This has been restored from {version}: {os.path.basename(path)}"
            )
        except (FileNotFoundError, OSError) as e:
            error_msg = f"ERROR: Failed to open Notepad++ at {NOTEPADPP_PATH}: {str(e)}"
            self.log_action(error_msg)
            QMessageBox.critical(self, "Error", f"Failed to open Notepad++: {str(e)}")

    def open_document_explorer(self):
        """Open the Document Version Explorer window."""
        try:
            self.doc_explorer = DocumentVersionExplorer()
            # Pass the current icon to the DocumentVersionExplorer
            current_icon = self.windowIcon()
            if not current_icon.isNull():
                self.doc_explorer.setWindowIcon(current_icon)
                # Also set the icon path in the DocumentVersionExplorer's config
                log_data = self.load_log()
                icon_path = log_data.get('icon_path')
                if icon_path:
                    self.doc_explorer.save_config('main_icon_path', icon_path)
                    self.doc_explorer.save_config('tray_icon_path', icon_path)
            self.doc_explorer.show()
            self.log_action("INFO: Opened Document Version Explorer")
        except Exception as e:
            error_msg = f"ERROR: Failed to open Document Version Explorer: {str(e)}"
            self.log_action(error_msg)
            QMessageBox.critical(self, "Error", f"Failed to open Document Version Explorer: {str(e)}")

    def close_full_application(self):
        """Close the entire application and all its windows."""
        self.log_action("INFO: Closing full application")
        # Close any open DocumentVersionExplorer windows
        if hasattr(self, 'doc_explorer') and self.doc_explorer is not None:
            self.doc_explorer.close()
        QApplication.quit()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    viewer = BackupViewer()
    viewer.show()
    sys.exit(app.exec())