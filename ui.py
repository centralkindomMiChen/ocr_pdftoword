import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QCheckBox, QTextEdit, QFileDialog, QMessageBox
)
from PyQt5.QtCore import QThread, pyqtSignal, Qt
import converter
import subprocess

class ConversionThread(QThread):
    progress_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool)

    def __init__(self, pdf_path, output_path, use_ocr, dpi, tesseract_path, poppler_path, tessdata_config):
        super().__init__()
        self.pdf_path = pdf_path
        self.output_path = output_path
        self.use_ocr = use_ocr
        self.dpi = dpi
        self.tesseract_path = tesseract_path
        self.poppler_path = poppler_path
        self.tessdata_config = tessdata_config
        self.doc = None

    def run(self):
        try:
            success = converter.convert_pdf_to_word(
                self.pdf_path, self.output_path, self.use_ocr, self.dpi,
                self.tesseract_path, self.poppler_path, self.tessdata_config,
                on_progress=lambda msg: self.progress_signal.emit(msg)
            )
            self.finished_signal.emit(success)
        except Exception as e:
            self.progress_signal.emit(f"Error in conversion thread: {str(e)}")
            self.finished_signal.emit(False)

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF to Word Converter")
        self.setGeometry(100, 100, 600, 400)

        # Enable translucency for Windows 7 Aero support
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setWindowOpacity(0.95)  # Slight transparency for gem-like effect

        layout = QVBoxLayout()

        # Input PDF
        input_layout = QHBoxLayout()
        self.input_label = QLabel("Input PDF:")
        self.input_path = QLineEdit()
        self.input_button = QPushButton("Browse")
        self.input_button.clicked.connect(self.browse_input)
        input_layout.addWidget(self.input_label)
        input_layout.addWidget(self.input_path)
        input_layout.addWidget(self.input_button)
        layout.addLayout(input_layout)

        # Output DOCX
        output_layout = QHBoxLayout()
        self.output_label = QLabel("Output DOCX:")
        self.output_path = QLineEdit()
        self.output_button = QPushButton("Browse")
        self.output_button.clicked.connect(self.browse_output)
        output_layout.addWidget(self.output_label)
        output_layout.addWidget(self.output_path)
        output_layout.addWidget(self.output_button)
        layout.addLayout(output_layout)

        # OCR Checkbox
        self.ocr_checkbox = QCheckBox("Enable OCR (for scanned PDFs)")
        self.ocr_checkbox.setChecked(True)
        layout.addWidget(self.ocr_checkbox)

        # Buttons
        buttons_layout = QHBoxLayout()
        self.start_button = QPushButton("Start Conversion")
        self.start_button.clicked.connect(self.start_conversion)
        self.stop_save_button = QPushButton("Stop and Save")
        self.stop_save_button.clicked.connect(lambda: self.stop_conversion(save=True))
        self.stop_save_button.setEnabled(False)
        self.stop_abandon_button = QPushButton("Stop and Abandon")
        self.stop_abandon_button.clicked.connect(lambda: self.stop_conversion(save=False))
        self.stop_abandon_button.setEnabled(False)
        self.open_button = QPushButton("Open Saved File")
        self.open_button.clicked.connect(self.open_output_file)
        self.open_button.setEnabled(False)
        buttons_layout.addWidget(self.start_button)
        buttons_layout.addWidget(self.stop_save_button)
        buttons_layout.addWidget(self.stop_abandon_button)
        buttons_layout.addWidget(self.open_button)
        layout.addLayout(buttons_layout)

        # Console Output
        self.console = QTextEdit()
        self.console.setReadOnly(True)
        layout.addWidget(self.console)

        self.setLayout(layout)

        # Apply gem-like blue styling with glossy effect
        self.setStyleSheet("""
            QWidget {
                background: qlineargradient(
                    x1: 0, y1: 0, x2: 1, y2: 1,
                    stop: 0 rgba(70, 130, 180, 215),
                    stop: 0.5 rgba(100, 149, 237, 215),
                    stop: 1 rgba(70, 130, 180, 215)
                );
                border: 1px solid rgba(255, 255, 255, 100);
                border-radius: 10px;
                color: white;
            }
            QLineEdit, QTextEdit {
                background: rgba(255, 255, 255, 200);
                border: 1px solid rgba(255, 255, 255, 150);
                border-radius: 5px;
                color: black;
            }
            QPushButton {
                background: qlineargradient(
                    x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 rgba(100, 149, 237, 255),
                    stop: 1 rgba(70, 130, 180, 255)
                );
                border: 1px solid rgba(255, 255, 255, 150);
                border-radius: 5px;
                padding: 5px;
                color: white;
            }
            QPushButton:hover {
                background: qlineargradient(
                    x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 rgba(135, 206, 250, 255),
                    stop: 1 rgba(100, 149, 237, 255)
                );
            }
            QPushButton:disabled {
                background: rgba(128, 128, 128, 200);
                color: rgba(255, 255, 255, 150);
            }
            QCheckBox, QLabel {
                background: transparent;
                color: white;
            }
        """)

        # Thread
        self.thread = None

    def browse_input(self):
        file = QFileDialog.getOpenFileName(self, "Select PDF", "", "PDF Files (*.pdf)")[0]
        if file:
            self.input_path.setText(file)
            output_dir = os.path.dirname(file)
            output_base = os.path.splitext(os.path.basename(file))[0]
            default_output = os.path.join(output_dir, f"convert_doc{output_base}.docx")
            self.output_path.setText(default_output)

    def browse_output(self):
        file = QFileDialog.getSaveFileName(self, "Select Output DOCX", "", "Word Files (*.docx)")[0]
        if file:
            self.output_path.setText(file)

    def start_conversion(self):
        pdf_path = self.input_path.text()
        output_path = self.output_path.text()
        use_ocr = self.ocr_checkbox.isChecked()

        if not pdf_path or not os.path.exists(pdf_path):
            QMessageBox.warning(self, "Error", "Invalid input PDF path.")
            return
        if not output_path:
            QMessageBox.warning(self, "Error", "Invalid output DOCX path.")
            return

        tesseract_path = r'D:\Python\Scripts\tesseract.exe'
        poppler_path = r'D:\Python\Scripts\poppler-24.02.0\Library\bin'
        tessdata_config = r'--tessdata-dir D:\Python\Scripts\tessdata'

        self.console.clear()
        self.start_button.setEnabled(False)
        self.stop_save_button.setEnabled(True)
        self.stop_abandon_button.setEnabled(True)
        self.open_button.setEnabled(False)

        self.thread = ConversionThread(pdf_path, output_path, use_ocr, 400, tesseract_path, poppler_path, tessdata_config)
        self.thread.progress_signal.connect(self.update_console)
        self.thread.finished_signal.connect(self.conversion_finished)
        self.thread.start()

    def stop_conversion(self, save=False):
        if self.thread:
            converter.cancel_conversion(save_partial=save, output_path=self.output_path.text())
            self.update_console("Stopping conversion...")

    def update_console(self, message):
        self.console.append(message)

    def conversion_finished(self, success):
        self.start_button.setEnabled(True)
        self.stop_save_button.setEnabled(False)
        self.stop_abandon_button.setEnabled(False)
        self.open_button.setEnabled(success)
        if success:
            self.update_console("Conversion successful!")
        else:
            self.update_console("Conversion failed or cancelled.")

    def open_output_file(self):
        output_path = self.output_path.text()
        if os.path.exists(output_path):
            try:
                if sys.platform == "win32":
                    os.startfile(output_path)
                elif sys.platform == "darwin":
                    subprocess.Popen(["open", output_path])
                else:
                    subprocess.Popen(["xdg-open", output_path])
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Could not open the file.\nError: {e}")
        else:
            QMessageBox.warning(self, "Error", "Output file does not exist.")

if __name__ == "__main__":
    try:
        app = QApplication(sys.argv)
        window = MainWindow()
        window.show()
        sys.exit(app.exec_())
    except Exception as e:
        print(f"Application error: {str(e)}")
