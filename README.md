# Web Scraper with Selenium and Python  

This project is a simple web scraper that uses **Selenium** to extract client information from a support website and saves the data to an **Excel file**.  

## Features  
- Automates web data extraction using Selenium  
- Supports dynamic URL changes for different client numbers  
- Saves scraped data into an Excel file  

## Requirements  
- Python 3.x  
- Google Chrome installed  
- Chrome WebDriver (`chromedriver.exe`)  
- Required Python libraries:  
  - `selenium`  
  - `pandas`  
  - `openpyxl`  
  - `webdriver-manager`  
## Installation  
1. Clone the repository:  
2. Install dependencies:
pip install -r requirements.txt
3. Download the correct version of ChromeDriver and place it in the project folder.

Usage
Edit client_numbers in the script to include the client numbers you want to scrape.
Run the script:
python scrape.py

The extracted data will be saved in output.xlsx.

Notes
Make sure your ChromeDriver version matches your Chrome browser version.
If SSL errors occur, ensure your Python libraries (urllib3, requests, certifi) are updated.

License
This project is open-source under the MIT License.

Let me know if you want any modifications! 🚀
