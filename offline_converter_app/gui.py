import json
import sys
from pathlib import Path
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QListWidget, QComboBox, QLineEdit,
    QCheckBox, QGroupBox, QFileDialog, QMessageBox, QProgressBar,
    QTextEdit, QFrame
)
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QIcon, QAction

from .common import ConversionOptions, ToolState, ensure_output_allowed, output_path_for
from .converter import convert_one_with_options
from .pdf_utils import compress_pdf, merge_pdfs, split_pdf

CONFIG_FILE = Path.home() / ".offline_converter" / "gui_settings.json"


def load_settings() -> dict:
    try:
        if CONFIG_FILE.exists():
            return json.loads(CONFIG_FILE.read_text())
    except Exception:
        pass
    return {}


def save_settings(settings: dict) -> None:
    try:
        CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
        CONFIG_FILE.write_text(json.dumps(settings))
    except Exception:
        pass


class ConverterWindow(QMainWindow):
    def __init__(self, tools: ToolState):
        super().__init__()
        self.tools = tools
        self.selected_files: list[Path] = []
        self.load_settings()
        self.init_ui()
    
    def load_settings(self) -> None:
        s = load_settings()
        self.output_dir = s.get("output_dir", str(Path.cwd()))
        self.target_format = s.get("target", "pdf")
        self.mode = s.get("mode", "auto")
        self.ocr = s.get("ocr", False)
        self.sensitive = s.get("sensitive", False)
        self.overwrite = s.get("overwrite", False)
        self.skip_existing = s.get("skip_existing", True)
        self.dark_mode = s.get("dark_mode", False)
    
    def save_settings(self) -> None:
        save_settings({
            "output_dir": self.output_dir,
            "target": self.target_format,
            "mode": self.mode,
            "ocr": self.ocr,
            "sensitive": self.sensitive,
            "overwrite": self.overwrite,
            "skip_existing": self.skip_existing,
            "dark_mode": self.dark_mode,
        })
    
    def init_ui(self) -> None:
        self.setWindowTitle("Offline Converter")
        self.setMinimumSize(700, 600)
        
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        
        title = QLabel("Offline Document Converter")
        title.setStyleSheet("font-size: 24px; font-weight: bold; padding: 10px;")
        layout.addWidget(title)
        
        file_group = QGroupBox("Input Files")
        file_layout = QVBoxLayout(file_group)
        
        file_btn_layout = QHBoxLayout()
        self.add_files_btn = QPushButton("Add Files")
        self.add_files_btn.clicked.connect(self.add_files)
        self.clear_btn = QPushButton("Clear")
        self.clear_btn.clicked.connect(self.clear_files)
        file_btn_layout.addWidget(self.add_files_btn)
        file_btn_layout.addWidget(self.clear_btn)
        file_btn_layout.addStretch()
        file_layout.addLayout(file_btn_layout)
        
        self.file_list = QListWidget()
        self.file_list.setMinimumHeight(100)
        file_layout.addWidget(self.file_list)
        layout.addWidget(file_group)
        
        settings_group = QGroupBox("Settings")
        settings_layout = QHBoxLayout(settings_group)
        
        settings_layout.addWidget(QLabel("Format:"))
        self.format_combo = QComboBox()
        self.format_combo.addItems(["pdf", "docx", "txt", "html", "md", "rtf", "pptx", "csv"])
        self.format_combo.setCurrentText(self.target_format)
        self.format_combo.currentTextChanged.connect(lambda t: setattr(self, 'target_format', t))
        settings_layout.addWidget(self.format_combo)
        
        settings_layout.addWidget(QLabel("Mode:"))
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["auto", "native", "python"])
        self.mode_combo.setCurrentText(self.mode)
        self.mode_combo.currentTextChanged.connect(lambda t: setattr(self, 'mode', t))
        settings_layout.addWidget(self.mode_combo)
        
        settings_layout.addStretch()
        layout.addWidget(settings_group)
        
        options_layout = QHBoxLayout()
        self.ocr_check = QCheckBox("OCR")
        self.ocr_check.setChecked(self.ocr)
        self.ocr_check.stateChanged.connect(lambda s: setattr(self, 'ocr', bool(s)))
        options_layout.addWidget(self.ocr_check)
        
        self.sensitive_check = QCheckBox("Sensitive Mode (300 DPI)")
        self.sensitive_check.setChecked(self.sensitive)
        self.sensitive_check.stateChanged.connect(lambda s: setattr(self, 'sensitive', bool(s)))
        options_layout.addWidget(self.sensitive_check)
        
        self.overwrite_check = QCheckBox("Overwrite")
        self.overwrite_check.setChecked(self.overwrite)
        self.overwrite_check.stateChanged.connect(lambda s: setattr(self, 'overwrite', bool(s)))
        options_layout.addWidget(self.overwrite_check)
        
        self.skip_check = QCheckBox("Skip Existing")
        self.skip_check.setChecked(self.skip_existing)
        self.skip_check.stateChanged.connect(lambda s: setattr(self, 'skip_existing', bool(s)))
        options_layout.addWidget(self.skip_check)
        
        options_layout.addStretch()
        layout.addLayout(options_layout)
        
        output_layout = QHBoxLayout()
        output_layout.addWidget(QLabel("Output:"))
        self.output_edit = QLineEdit(self.output_dir)
        self.output_edit.textChanged.connect(lambda t: setattr(self, 'output_dir', t))
        output_layout.addWidget(self.output_edit, 1)
        
        self.browse_btn = QPushButton("Browse")
        self.browse_btn.clicked.connect(self.browse_output)
        output_layout.addWidget(self.browse_btn)
        layout.addLayout(output_layout)
        
        self.convert_btn = QPushButton("Convert Files")
        self.convert_btn.setMinimumHeight(50)
        self.convert_btn.setStyleSheet("background-color: #16a34a; color: white; font-size: 14px; font-weight: bold;")
        self.convert_btn.clicked.connect(self.run_convert)
        layout.addWidget(self.convert_btn)
        
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        layout.addWidget(self.progress)
        
        results_group = QGroupBox("Results")
        results_layout = QVBoxLayout(results_group)
        self.results_text = QTextEdit()
        self.results_text.setReadOnly(True)
        self.results_text.setMaximumHeight(150)
        results_layout.addWidget(self.results_text)
        layout.addWidget(results_group)
        
        util_layout = QHBoxLayout()
        self.merge_btn = QPushButton("Merge PDFs")
        self.merge_btn.clicked.connect(self.merge_pdfs)
        util_layout.addWidget(self.merge_btn)
        
        self.split_btn = QPushButton("Split PDF")
        self.split_btn.clicked.connect(self.split_pdf)
        util_layout.addWidget(self.split_btn)
        
        self.compress_btn = QPushButton("Compress PDF")
        self.compress_btn.clicked.connect(self.compress_pdf)
        util_layout.addWidget(self.compress_btn)
        
        util_layout.addStretch()
        layout.addLayout(util_layout)
        
        self.statusBar().showMessage("Ready")
    
    def add_files(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(self, "Select Files", "", "All Files (*)")
        for f in files:
            p = Path(f)
            if p not in self.selected_files:
                self.selected_files.append(p)
                self.file_list.addItem(p.name)
    
    def clear_files(self) -> None:
        self.selected_files.clear()
        self.file_list.clear()
    
    def browse_output(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder", self.output_dir)
        if folder:
            self.output_dir = folder
            self.output_edit.setText(folder)
    
    def run_convert(self) -> None:
        if not self.selected_files:
            QMessageBox.warning(self, "No Files", "Please add files to convert.")
            return
        
        self.save_settings()
        self.convert_btn.setEnabled(False)
        self.progress.setVisible(True)
        self.progress.setMaximum(len(self.selected_files))
        self.results_text.clear()
        
        output_dir = Path(self.output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        
        success = 0
        failed = 0
        
        for idx, input_path in enumerate(self.selected_files):
            self.statusBar().showMessage(f"Converting: {input_path.name}")
            self.progress.setValue(idx + 1)
            QApplication.processEvents()
            
            dest = output_dir / f"{input_path.stem}.{self.target_format}"
            
            if dest.exists() and self.skip_existing:
                self.results_text.append(f"Skipped: {dest.name}")
                continue
            
            if dest.exists() and not self.overwrite:
                self.results_text.append(f"Exists: {dest.name}")
                failed += 1
                continue
            
            try:
                report = convert_one_with_options(
                    input_path,
                    self.target_format,
                    dest,
                    self.tools,
                    ConversionOptions(
                        ocr=self.ocr,
                        mode=self.mode,
                        sensitive=self.sensitive,
                    )
                )
                self.results_text.append(f"✓ {dest.name} - {report.engine}")
                success += 1
            except Exception as e:
                self.results_text.append(f"✗ {input_path.name}: {e}")
                failed += 1
        
        self.statusBar().showMessage(f"Done: {success} converted, {failed} failed")
        self.progress.setVisible(False)
        self.convert_btn.setEnabled(True)
        self.convert_btn.setText("✓ Convert Complete")
        QTimer.singleShot(3000, lambda: self.convert_btn.setText("Convert Files"))
    
    def merge_pdfs(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(self, "Select PDFs to Merge", "", "PDF Files (*.pdf)")
        if len(files) < 2:
            return
        
        output, _ = QFileDialog.getSaveFileName(self, "Save Merged PDF", "", "PDF Files (*.pdf)")
        if not output:
            return
        
        try:
            merge_pdfs([Path(f) for f in files], Path(output))
            QMessageBox.information(self, "Done", f"Merged to:\n{output}")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))
    
    def split_pdf(self) -> None:
        file, _ = QFileDialog.getOpenFileName(self, "Select PDF to Split", "", "PDF Files (*.pdf)")
        if not file:
            return
        
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if not folder:
            return
        
        try:
            split_pdf(Path(file), Path(folder))
            QMessageBox.information(self, "Done", f"Split to:\n{folder}")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))
    
    def compress_pdf(self) -> None:
        file, _ = QFileDialog.getOpenFileName(self, "Select PDF to Compress", "", "PDF Files (*.pdf)")
        if not file:
            return
        
        output, _ = QFileDialog.getSaveFileName(self, "Save Compressed PDF", "", "PDF Files (*.pdf)")
        if not output:
            return
        
        try:
            compress_pdf(Path(file), Path(output))
            QMessageBox.information(self, "Done", f"Compressed to:\n{output}")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))


def launch_gui(tools: ToolState) -> int:
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = ConverterWindow(tools)
    window.show()
    return app.exec()