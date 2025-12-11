import time
import os
import cv2
import numpy as np
import pytesseract
from pytesseract import Output
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.colors import Color
from pypdf import PdfWriter, PdfReader
from io import BytesIO

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.service import Service
from webdriver_manager.firefox import GeckoDriverManager

# --- CONFIGURATION ---
PREZI_URL = "" # URL of the Prezi presentation
TOTAL_SLIDES = 1  # You must count the steps manually beforehand
WAIT_TIME = 4      # Seconds to wait for animation (increase if internet is slow)
OUTPUT_PDF = "" # Output PDF file name

# Cropping
CROP_TOP = 80            
CROP_BOTTOM = 80         

# Tesseract Path (Windows Only) - Uncomment if needed
pytesseract.pytesseract.tesseract_cmd = r'S:\Program Files\Tesseract-OCR\tesseract.exe'
# ---------------------

def get_clean_image_for_ocr(pil_image):
    """
    Creates a high-contrast binary version of the image solely for OCR detection.
    """
    open_cv_image = np.array(pil_image) 
    gray = cv2.cvtColor(open_cv_image, cv2.COLOR_RGB2GRAY)
    
    # Adaptive Thresholding to fix the gradient background issue
    binary = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
        cv2.THRESH_BINARY_INV, 31, 15
    )
    return Image.fromarray(binary)

def create_overlay_pdf(original_img, ocr_img):
    """
    Draws the ORIGINAL image on a PDF page, then detects text from the OCR image
    and writes it invisibly on top.
    """
    img_width, img_height = original_img.size
    
    # 1. Get text data from the 'Ugly' OCR image
    # We use a configuration that treats the image as a single block of text (psm 11 or 3) if needed,
    # but default usually works for word coordinates.
    data = pytesseract.image_to_data(ocr_img, output_type=Output.DICT)
    
    # 2. Setup PDF Canvas
    pdf_buffer = BytesIO()
    c = canvas.Canvas(pdf_buffer, pagesize=(img_width, img_height))
    
    # 3. Draw the ORIGINAL (Colorful) image as the background
    # We save the PIL image to bytes so ReportLab can read it
    img_buffer = BytesIO()
    original_img.save(img_buffer, format='PNG')
    img_buffer.seek(0)
    c.drawImage(
        flask_image_reader(img_buffer), # Helper defined below or simply filename
        0, 0, width=img_width, height=img_height
    )

    # 4. Draw Invisible Text over it
    c.setFillColor(Color(0, 0, 0, alpha=0)) # Fully transparent ink
    # c.setFillColor(Color(1, 0, 0, alpha=0.5)) # DEBUG: Uncomment to see red text overlay
    
    n_boxes = len(data['text'])
    for i in range(n_boxes):
        text = data['text'][i].strip()
        if not text:
            continue
            
        # Get coordinates
        x, y, w, h = data['left'][i], data['top'][i], data['width'][i], data['height'][i]
        
        # ReportLab coordinate system (0,0) is Bottom-Left. 
        # Tesseract (0,0) is Top-Left. We must flip Y.
        pdf_y = img_height - y - h
        
        # Draw the text
        c.setFont("Helvetica", h) # Approximate font size by box height
        c.drawString(x, pdf_y, text)

    c.save()
    pdf_buffer.seek(0)
    return pdf_buffer

# Helper for ReportLab to read image from memory
from reportlab.lib.utils import ImageReader
def flask_image_reader(img_buffer):
    return ImageReader(img_buffer)


def capture_prezi():
    options = webdriver.FirefoxOptions()
    # options.add_argument("--headless") 
    
    driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()), options=options)
    final_pdf_writer = PdfWriter()
    temp_files = []

    try:
        print(f"Opening: {PREZI_URL}")
        driver.get(PREZI_URL)
        print("Waiting for Prezi to load (15s)...")
        time.sleep(15) 
        body = driver.find_element(By.TAG_NAME, "body")

        for i in range(TOTAL_SLIDES):
            print(f"Processing slide {i+1}/{TOTAL_SLIDES}...")
            
            # 1. Capture
            temp_filename = f"temp_slide_{i}.png"
            driver.save_screenshot(temp_filename)
            temp_files.append(temp_filename)
            
            with Image.open(temp_filename) as img:
                width, height = img.size
                
                # 2. Crop
                crop_box = (0, CROP_TOP, width, height - CROP_BOTTOM)
                cropped_img = img.crop(crop_box)
                
                # 3. Create 'Ugly' High-Contrast copy for OCR
                ocr_img = get_clean_image_for_ocr(cropped_img)
                
                # 4. Generate PDF Page (Original Visuals + Hidden Text)
                page_pdf_bytes = create_overlay_pdf(cropped_img, ocr_img)
                
                # 5. Add to final PDF
                page_reader = PdfReader(page_pdf_bytes)
                final_pdf_writer.add_page(page_reader.pages[0])
            
            body.send_keys(Keys.ARROW_RIGHT)
            time.sleep(WAIT_TIME)

        print("Capture complete. Saving final PDF...")
        with open(OUTPUT_PDF, "wb") as f:
            final_pdf_writer.write(f)
        print(f"Success! Saved to {OUTPUT_PDF}")

    except Exception as e:
        print(f"An error occurred: {e}")
        import traceback
        traceback.print_exc()
        
    finally:
        print("Cleaning up...")
        driver.quit()
        for file in temp_files:
            try:
                if os.path.exists(file):
                    os.remove(file)
            except Exception:
                pass

if __name__ == "__main__":
    capture_prezi()