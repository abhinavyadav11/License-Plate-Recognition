import time
import pytesseract
import cv2
import numpy as np
from PIL import Image
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# Set Chrome options for headless execution
chrome_options = Options()
chrome_options.add_argument("--headless")  # Run without GUI
chrome_options.add_argument("--no-sandbox")  # Required for cloud instances
chrome_options.add_argument("--disable-dev-shm-usage")  # Prevent crashes
chrome_options.add_argument("--user-data-dir=/tmp/selenium")  # Avoids conflicts

# Initialize WebDriver
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

# List to store extracted data
data = []

try:
    # -------------------------------
    # Step 1: Open website and submit search form
    # -------------------------------
    driver.get("https://iprsearch.ipindia.gov.in/PublicSearch")
    
    # Locate date fields and input values
    start_date = driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/div[4]/form/section/div/div/div/div/div[2]/div[2]/div/div[2]/input")
    end_date = driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/div[4]/form/section/div/div/div/div/div[2]/div[4]/div/div/input")
    start_date.send_keys("01-01-2022")
    end_date.send_keys("12-31-2022")
    
    # -------------------------------
    # Step 2: Capture and Process CAPTCHA
    # -------------------------------
    captcha_element = driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/div[4]/form/section/div/div/div/div/div[18]/div[2]/div/div/div[1]/div[1]/img[1]")
    captcha_element.screenshot("captcha.png")
    print("CAPTCHA saved as 'captcha.png'. Attempting OCR...")


    # Load image and preprocess
    img = cv2.imread("captcha.png", cv2.IMREAD_GRAYSCALE)
    _, img = cv2.threshold(img, 150, 255, cv2.THRESH_BINARY)  # Thresholding for better accuracy
    cv2.imwrite("processed_captcha.png", img)


    # Perform OCR
    captcha_text = pytesseract.image_to_string(Image.open("processed_captcha.png"), config="--psm 6").strip()
    print(f"Extracted CAPTCHA: {captcha_text}")

    # If OCR fails, ask for manual input
    if len(captcha_text) < 3:
        captcha_text = input("OCR failed. Enter CAPTCHA manually: ")

    # Enter CAPTCHA into the form
    captcha_input = driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/div[4]/form/section/div/div/div/div/div[18]/div[2]/div/div/div[1]/div[2]/input")
    captcha_input.send_keys(captcha_text)

    # Submit the form
    search_button = driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/div/div[4]/form/section/div/div/div/div/div[18]/div[2]/div/div/div[1]/div[3]/input")
    search_button.click()
    print("Form submitted successfully!")

    # -------------------------------
    # Step 3: Extract Search Results
    # -------------------------------
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//table/tbody/tr[1]/td[1]/button"))
    )

    main_window = driver.window_handles[0]
    page_number = 1
    
    while True:
        print(f"\nProcessing page {page_number}...")
        app_rows = driver.find_elements(By.XPATH, "//table/tbody/tr")
        total_apps = len(app_rows)
        print(f"Found {total_apps} applications on current page.")

        for idx in range(total_apps):
            app_rows = driver.find_elements(By.XPATH, "//table/tbody/tr")
            current_row = app_rows[idx]

            # Extract Application Number and Title
            cells = current_row.find_elements(By.XPATH, "./td")
            if len(cells) < 2:
                print(f"Skipping row {idx+1}: Not enough columns.")
                continue

            app_number = cells[0].text.strip()
            title = cells[1].text.strip()
            print(f"Processing Application: {app_number}, Title: {title}")

            # Click the application link
            current_row.find_element(By.XPATH, "./td[1]/button").click()
            time.sleep(2)

            # Switch to new window
            WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > 1)
            driver.switch_to.window(driver.window_handles[1])
            print("Switched to application details window.")

            # Extract complete specification
            spec_xpath = "/html/body/div[1]/div[2]/div/div[4]/form/div/div/table/tbody/tr[17]/td/textarea"
            try:
                complete_spec_element = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, spec_xpath))
                )
                complete_spec = complete_spec_element.text.strip()
                print(f"Extracted Specification for {app_number}:\n{complete_spec}\n")
            except:
                complete_spec = "Not Available"
                print(f"No specification found for {app_number}. Skipping...\n")

            data.append({
                "Application Number": app_number,
                "Title": title,
                "Specification": complete_spec
            })

            # Close the application details window
            driver.close()
            print("Closed application details window.")
            driver.switch_to.window(main_window)
            time.sleep(2)

        # Pagination Handling
        try:
            next_button_xpath = "/html/body/div[1]/div[2]/div/div[4]/div/div[2]/table/tfoot/tr/td/table/tbody/tr/th[2]/button[3]"
            next_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, next_button_xpath))
            )
            if "disabled" in next_button.get_attribute("class"):
                print("No more pages to process.")
                break
            next_button.click()
            print("Clicked Next button; loading next page...")
            time.sleep(4)
            page_number += 1
        except Exception as e:
            print(f"Error clicking Next button: {e}")
            break

    print("All pages processed successfully.")

finally:
    driver.quit()
    print("Browser closed, process completed successfully!")
