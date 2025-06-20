from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.keys import Keys
import time
import os
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import requests

# Path to msedgedriver.exe (assumed to be in the same directory)
driver_path = os.path.join(os.path.dirname(__file__), 'msedgedriver.exe')

# Set up Edge options to download PDFs automatically
options = webdriver.EdgeOptions()
download_dir = r"Z:/Coding eh/ResultReaper/Results"
options.add_experimental_option("prefs", {
    "download.default_directory": download_dir,
    "plugins.always_open_pdf_externally": True,  # Download PDFs instead of opening
    "download.prompt_for_download": False,
    "download.directory_upgrade": True
})

# Set up Edge WebDriver
service = Service(driver_path)
driver = webdriver.Edge(service=service, options=options)

# Open the results page
driver.get('https://siuexam.siu.edu.in/forms/resultview.html')

# Wait for the PRN input to be present (wait up to 10 seconds)
wait = WebDriverWait(driver, 10)
prn_input = wait.until(
    EC.presence_of_element_located((By.ID, "login"))
)

# Enter the hardcoded PRN
prn = '23000000000'
prn_input.clear()
prn_input.send_keys(prn)

# Click the Login button (wait for it to be clickable)
login_button = wait.until(
    EC.element_to_be_clickable((By.ID, "lgnbtn"))
)
login_button.click()

# Wait for the seat number input to appear
seat_input = wait.until(
    EC.presence_of_element_located((By.ID, "txt4"))
)
seat_input.clear()

# change the following to the acutual seat number
seat_input.send_keys("50000")


# Click the View button
view_button = wait.until(
    EC.element_to_be_clickable((By.XPATH, "//input[@type='button' and @value='View' and contains(@class, 'btn-info')]"))
)
view_button.click()

# Wait for the 'Click to Download' button to appear
# The button is an <a> tag with class 'btn-success btndef' and text 'Click to Download'
download_button = wait.until(
    EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'btn-success') and contains(@class, 'btndef') and contains(text(), 'Click to Download')]"))
)
pdf_url = download_button.get_attribute('href')

if pdf_url:
    # Build the absolute URL if needed
    if pdf_url.startswith('..'):
        from urllib.parse import urljoin
        pdf_url = urljoin(driver.current_url, pdf_url)

    # Download the PDF using requests
    pdf_response = requests.get(pdf_url)
    pdf_filename = os.path.join(download_dir, f'{prn}.pdf')
    with open(pdf_filename, 'wb') as f:
        f.write(pdf_response.content)
    print(f"PDF downloaded to: {pdf_filename}")
else:
    print("Could not find the PDF download link.")

input("Press Enter to exit and close the browser...")
