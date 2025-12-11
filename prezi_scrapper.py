import time
import img2pdf
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
# These imports handle the Firefox driver automatically
from selenium.webdriver.firefox.service import Service
from webdriver_manager.firefox import GeckoDriverManager

# --- CONFIGURATION ---
PREZI_URL = "Slide URL here"  # Replace with your Prezi URL
TOTAL_SLIDES = 24  # You must count the steps manually beforehand
WAIT_TIME = 4      # Seconds to wait for animation (increase if internet is slow)
OUTPUT_PDF = "prezi_presentation.pdf"
# ---------------------

def capture_prezi():
    # 1. Setup Firefox
    options = webdriver.FirefoxOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--window-size=1920,1080")
    # options.add_argument("--headless") # Uncomment to run invisible
    
    print("Setting up Firefox Driver...")
    # This automatically downloads and links the correct geckodriver
    driver = webdriver.Firefox(service=Service(GeckoDriverManager().install()), options=options)
    
    try:
        print(f"Opening: {PREZI_URL}")
        driver.get(PREZI_URL)
        
        # 2. Initial Load Wait
        print("Waiting for Prezi to load (15s)...")
        time.sleep(15) 
        
        # Locate the body to send keystrokes
        body = driver.find_element(By.TAG_NAME, "body")
        
        image_files = []

        # 3. Loop through slides
        for i in range(TOTAL_SLIDES):
            print(f"Capturing slide {i+1}/{TOTAL_SLIDES}...")
            
            # Save screenshot
            filename = f"slide_{i:03d}.png"
            driver.save_screenshot(filename)
            image_files.append(filename)
            
            # Navigate Next
            body.send_keys(Keys.ARROW_RIGHT)
            
            # Wait for animation
            time.sleep(WAIT_TIME)

        print("Capture complete. Generating PDF...")
        
        # 4. Convert to PDF
        # Note: img2pdf requires the image files list
        with open(OUTPUT_PDF, "wb") as f:
            f.write(img2pdf.convert(image_files))
            
        print(f"Success! Saved to {OUTPUT_PDF}")

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    capture_prezi()