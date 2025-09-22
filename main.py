import os
import pytesseract
from converter import convert_pdf_to_word
from ui import MainWindow
from PyQt5.QtWidgets import QApplication

def run_command_line():
    """
    Run the conversion from command line with hardcoded paths.
    Sets default output path based on input PDF.
    """
    TESSERACT_PATH = r'D:\Python\Scripts\tesseract.exe'
    POPPLER_PATH = r'D:\Python\Scripts\poppler-24.02.0\Library\bin'
    INPUT_PDF_PATH = r'K:\关于修订《中国国际航空股份有限公司ICS订座系统工作号管理与使用规定》的通知.pdf'
    TESSDATA_CONFIG = r'--tessdata-dir D:\Python\Scripts\tessdata'

    # Set default output path
    output_dir = os.path.dirname(INPUT_PDF_PATH)
    output_base = os.path.splitext(os.path.basename(INPUT_PDF_PATH))[0]
    OUTPUT_DOCX_PATH = os.path.join(output_dir, f"convert_doc{output_base}.docx")

    if __name__ == "__main__":
        try:
            pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH
            success = convert_pdf_to_word(INPUT_PDF_PATH, OUTPUT_DOCX_PATH, use_ocr=True, dpi=400,
                                        tesseract_path=TESSERACT_PATH, poppler_path=POPPLER_PATH,
                                        tessdata_config=TESSDATA_CONFIG, on_progress=print)
            if success:
                print("Conversion completed successfully.")
            else:
                print("Conversion failed or was cancelled.")
        except Exception as e:
            print(f"Command-line error: {str(e)}")

def run_ui():
    """
    Run the application with the graphical user interface.
    """
    try:
        app = QApplication([])
        window = MainWindow()
        window.show()
        app.exec_()
    except Exception as e:
        print(f"UI error: {str(e)}")

if __name__ == "__main__":
    # Choose to run either the command-line version or UI
    # Uncomment the desired mode:
    run_ui()  # Run with GUI
    # run_command_line()  # Run with command line
