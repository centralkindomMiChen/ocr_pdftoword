import os
from PIL import Image
import pytesseract
from pdf2image import convert_from_path
from docx import Document
import cv2
import numpy as np
import threading
from pdf2docx import Converter
import fitz  # PyMuPDF for getting page count

# Global flag for cancellation
cancel_flag = threading.Event()

def preprocess_image(image):
    """
    Preprocess the image to improve OCR accuracy with a lighter approach.
    - Convert to grayscale
    - Apply basic thresholding
    """
    try:
        open_cv_image = np.array(image.convert('RGB'))[:, :, ::-1].copy()
        gray = cv2.cvtColor(open_cv_image, cv2.COLOR_BGR2GRAY)
        
        # Simple binary thresholding
        _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
        
        processed_image = Image.fromarray(thresh)
    except ImportError:
        print("OpenCV not found, using basic PIL preprocessing.")
        processed_image = image.convert('L')
        processed_image = processed_image.point(lambda x: 0 if x < 150 else 255, '1')
    return processed_image

def convert_pdf_to_word(pdf_path, output_path, use_ocr=True, dpi=400, tesseract_path=None, poppler_path=None, tessdata_config=None, on_progress=None):
    """
    Convert PDF to Word.
    - If OCR is disabled (use_ocr=False), it uses the pdf2docx library for a direct conversion.
    - If OCR is enabled (use_ocr=True), it performs OCR on each page.
    - Supports cancellation via global cancel_flag.
    """
    global cancel_flag
    cancel_flag.clear()

    if on_progress:
        on_progress("Starting PDF to Word conversion...")

    if not use_ocr:
        # --- BUG FIX ---
        # The original page-by-page merging created corrupted files.
        # This is the correct, simplified approach using pdf2docx directly.
        try:
            if on_progress:
                on_progress("Using pdf2docx library for direct conversion...")

            # Initialize the converter
            cv = Converter(pdf_path)
            
            # Convert the entire document. This is the correct usage.
            # The `convert` method handles all pages by default.
            cv.convert(output_path, start=0, end=None)
            
            # Close the converter to free up resources
            cv.close()

            # Check for cancellation flag after the main blocking operation
            if cancel_flag.is_set():
                if on_progress: on_progress("Conversion was cancelled after completion. The file was saved.")
                return False # Or True, depending on desired behavior

            if on_progress:
                on_progress(f"Conversion completed successfully. Output saved to: {output_path}")
            return True
        except Exception as e:
            if on_progress:
                on_progress(f"Error in pdf2docx conversion: {str(e)}.")
            return False
        # --- END OF FIX ---

    # OCR conversion logic remains the same
    doc = Document()

    if tesseract_path:
        pytesseract.pytesseract.tesseract_cmd = tesseract_path

    try:
        images = convert_from_path(pdf_path, dpi=dpi, poppler_path=poppler_path)
    except Exception as e:
        if on_progress:
            on_progress(f"Error converting PDF to images: {str(e)}")
        return False

    num_images = len(images)
    for i, image in enumerate(images):
        if cancel_flag.is_set():
            if on_progress:
                on_progress("Conversion cancelled.")
            # If partial saving is desired during OCR, it could be implemented here.
            return False

        if on_progress:
            on_progress(f"Processing page {i+1}/{num_images} via OCR...")

        processed_image = preprocess_image(image)
        config = tessdata_config + ' --psm 6 --oem 1' if tessdata_config else '--psm 6 --oem 1'
        text = pytesseract.image_to_string(processed_image, lang='chi_sim', config=config)
        
        doc.add_paragraph(text)

        if i < num_images - 1:
            doc.add_page_break()

    if cancel_flag.is_set():
        if on_progress:
            on_progress("Conversion cancelled before saving.")
        return False

    doc.save(output_path)
    if on_progress:
        on_progress(f"Conversion completed. Output saved to: {output_path}")

    return True

def cancel_conversion(save_partial=False, output_path=None, doc=None):
    """
    Sets a flag to cancel the ongoing conversion.
    Note: For the direct pdf2docx conversion, cancellation is not granular and will
    stop the process before it starts or after it finishes. For OCR, it's page-by-page.
    """
    global cancel_flag
    cancel_flag.set()
    # Partial saving for OCR is complex and not fully implemented here.
    # The original implementation was also flawed as 'doc' was local.
    if save_partial and doc and output_path:
        print("Note: Partial saving is only effective for OCR mode.")
        doc.save(output_path)
