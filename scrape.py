from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
import os
import pandas as pd

# Set the correct WebDriver path
WEBDRIVER_PATH = os.path.expanduser("~/desktop/chromedriver.exe")

# Set up Chrome options
chrome_options = Options()
chrome_options.add_argument("--headless")  # Run in background
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--ignore-certificate-errors")
chrome_options.add_argument("--allow-running-insecure-content")
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

# Initialize WebDriver
service = Service(WEBDRIVER_PATH)
driver = webdriver.Chrome(service=service, options=chrome_options)

def get_client_info(client_number):
    """Fetches client details from the support page."""
    url = f"https://support.meditab.com/proxysupport/support/?i={client_number}"
    driver.get(url)
    time.sleep(2)  # Reduce load time to avoid being blocked

    try:
        email = driver.find_element(By.XPATH, "//input[@id='txt_email']").get_attribute("value")
    except:
        email = "N/A"

    try:
        office = driver.find_element(By.XPATH, "//input[@id='CustomField1']").get_attribute("value")
    except:
        office = "N/A"

    try:
        contact = driver.find_element(By.XPATH, "//input[@id='txt_contact']").get_attribute("value")
    except:
        contact = "N/A"

    return {
        "Client Number": client_number,
        "Email": email,
        "Office": office,
        "Contact": contact
    }

# 🔹 Generate 100 client numbers (or replace with your own list)
client_numbers = [str(50329 + i).zfill(7) for i in range(1000)]

# Store results
client_data = []

for index, client in enumerate(client_numbers, start=1):
    print(f"Fetching data for {client} ({index}/1000)...")
    try:
        client_info = get_client_info(client)
        client_data.append(client_info)
        print(f"✅ {client_info}\n")
    except Exception as e:
        print(f"❌ Error fetching {client}: {e}")
    time.sleep(1)  # Add delay to prevent detection

# Convert list to DataFrame
df = pd.DataFrame(client_data)

# Save to Excel file
output_file = os.path.expanduser("~/Desktop/client_info.xlsx")
try:
    df.to_excel(output_file, index=False)
    print(f"✅ Excel file saved at: {output_file}")
except Exception as e:
    print(f"❌ Error saving Excel file: {e}")

# Quit WebDriver
driver.quit()
print("✅ Done!")
