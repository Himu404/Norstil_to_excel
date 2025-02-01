# Company Profile Scraping

This Python script scrapes company profiles from the Nordstil Messe Frankfurt website using Selenium. It extracts details like company name, address, contact info, and website, and saves them into an Excel file for easy analysis.

## Features
- Scrapes company details such as name, address, phone, email, and website.
- Saves the extracted data to an Excel file.
- Handles popups and retries on failed extractions.
- Automatically navigates through multiple pages to scrape all profiles.

## Requirements
- Python 3.x
- Selenium
- OpenPyXL
- Chrome WebDriver (ensure the correct path to `chromedriver` is set)

## Installation

1. Install the required Python libraries:
   ```bash
   pip install selenium openpyxl

## Usage

1. Clone the repository or download the script.
2. Set up the `chrome_driver_path` in the script to your local `chromedriver` location.
3. Run the script:
   ```bash
   python scrape_profiles.py
