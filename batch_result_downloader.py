import csv
import os
import time
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

MAIN_CSV = 'main.csv'
MAPPING_CSV = 'mappings.csv'
RESULTS_DIR = r"Z:/Coding eh/ResultReaper/Results" #change this to the desired directory 
DRIVER_PATH = os.path.join(os.path.dirname(__file__), 'msedgedriver.exe')
RESULT_URL = 'https://siuexam.siu.edu.in/forms/resultview.html'

# Load main.csv into a list of dicts
def load_main():
    with open(MAIN_CSV, newline='', encoding='utf-8') as f:
        return list(csv.DictReader(f))

def save_main(rows):
    with open(MAIN_CSV, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=rows[0].keys())
        writer.writeheader()
        writer.writerows(rows)

# Load mappings.csv into a set of (prn, seat_no)
def load_mappings():
    if not os.path.exists(MAPPING_CSV):
        return set()
    with open(MAPPING_CSV, newline='', encoding='utf-8') as f:
        return set((row['mapped_prn'], row['mapped_seat_no']) for row in csv.DictReader(f))

def append_mapping(prn, seat_no):
    file_exists = os.path.exists(MAPPING_CSV)
    with open(MAPPING_CSV, 'a', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=['mapped_prn', 'mapped_seat_no'])
        if not file_exists:
            writer.writeheader()
        writer.writerow({'mapped_prn': prn, 'mapped_seat_no': seat_no})

# Ensure results directory exists
os.makedirs(RESULTS_DIR, exist_ok=True)

def main():
    main_rows = load_main()
    mappings = load_mappings()
    used_seats = set(seat_no for _, seat_no in mappings)
    prn_to_row = {row['prn']: row for row in main_rows}

    for row in main_rows:
        prn = row['prn']
        if row['prn_status'] == 'downloaded':
            continue
        print(f"Processing PRN: {prn}")

        # Set up Edge WebDriver for each PRN
        options = webdriver.EdgeOptions()
        options.add_experimental_option("prefs", {
            "download.default_directory": RESULTS_DIR,
            "plugins.always_open_pdf_externally": True,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True
        })
        service = Service(DRIVER_PATH)
        driver = webdriver.Edge(service=service, options=options)
        wait = WebDriverWait(driver, 10)
        try:
            driver.get(RESULT_URL)
            # Enter PRN
            prn_input = wait.until(EC.presence_of_element_located((By.ID, "login")))
            prn_input.clear()
            prn_input.send_keys(prn)
            login_button = wait.until(EC.element_to_be_clickable((By.ID, "lgnbtn")))
            login_button.click()
            # Wait for seat number input or error
            try:
                seat_input = wait.until(EC.presence_of_element_located((By.ID, "txt4")))
            except:
                print(f"PRN {prn} not valid or result not declared.")
                row['prn_status'] = 'error'
                save_main(main_rows)
                driver.quit()
                continue
            # Try all unused seat numbers
            for seat_row in main_rows:
                seat_no = seat_row['seat_no']
                if seat_row['seat_status'] == 'used':
                    continue
                # Try seat number
                seat_input.clear()
                seat_input.send_keys(seat_no)
                view_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='button' and @value='View' and contains(@class, 'btn-info')]")))
                view_button.click()
                time.sleep(1.5)  # Wait for response
                # Check for download button
                try:
                    download_button = driver.find_element(By.XPATH, "//a[contains(@class, 'btn-success') and contains(@class, 'btndef') and contains(text(), 'Click to Download')]")
                    pdf_url = download_button.get_attribute('href')
                    if pdf_url:
                        from urllib.parse import urljoin
                        pdf_url = urljoin(driver.current_url, pdf_url)
                        pdf_response = requests.get(pdf_url)
                        pdf_filename = os.path.join(RESULTS_DIR, f'{prn}.pdf')
                        with open(pdf_filename, 'wb') as f:
                            f.write(pdf_response.content)
                        print(f"Downloaded result for PRN {prn} with seat number {seat_no}")
                        row['prn_status'] = 'downloaded'
                        seat_row['seat_status'] = 'used'
                        append_mapping(prn, seat_no)
                        save_main(main_rows)
                        driver.quit()
                        break
                except:
                    # Check for error message
                    try:
                        error_msg = driver.find_element(By.XPATH, "//*[contains(text(), 'Enter valid seat no')]")
                        if error_msg.is_displayed():
                            continue
                    except:
                        continue
            else:
                # No seat number worked
                row['prn_status'] = 'error'
                save_main(main_rows)
                driver.quit()
        except Exception as e:
            print(f"Error processing PRN {prn}: {e}")
            row['prn_status'] = 'error'
            save_main(main_rows)
            driver.quit()

if __name__ == "__main__":
    main() 