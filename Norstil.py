import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl  # To work with Excel files

# Initialize Chrome WebDriver
chrome_driver_path = 'C:/Users/Himu/Desktop/PY Projects/chromedriver-win64/chromedriver.exe'
service = Service(chrome_driver_path)
options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)
driver = webdriver.Chrome(service=service, options=options)

# Set to keep track of visited profile URLs
visited_profiles = set()

# Function to extract company data with retry logic
def extract_profile_data(retries=3):
    data = {}
    attempt = 0
    while attempt < retries:
        try:
            # Wait until the address container is loaded
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".ex-contact-box__address-field-full-address")))

            company_name = driver.find_element(By.CSS_SELECTOR, ".ex-contact-box__address-field-full-address").text.split("\n")[0]
            street_address = driver.find_element(By.CSS_SELECTOR, ".ex-contact-box__address-field-full-address").text.split("\n")[1] if len(driver.find_element(By.CSS_SELECTOR, ".ex-contact-box__address-field-full-address").text.split("\n")) > 1 else ""
            
            # Now handle the third and fourth lines (ZIP/City and Country)
            city_zip_country = driver.find_element(By.CSS_SELECTOR, ".ex-contact-box__address-field-full-address").text.split("\n")[2:]  # Get the third and the rest
            zip_city = city_zip_country[0] if len(city_zip_country) > 0 else ""
            country = city_zip_country[-1] if len(city_zip_country) > 1 else ""

            # Split zip_city into zip and city
            zip_code, city = "", ""
            if zip_city:
                parts = zip_city.split(" ", 1)
                if len(parts) == 2:
                    zip_code, city = parts

            telephone = driver.find_element(By.CSS_SELECTOR, ".ex-contact-box__address-field-tel-number").text if driver.find_elements(By.CSS_SELECTOR, ".ex-contact-box__address-field-tel-number") else ""
            fax = driver.find_element(By.CSS_SELECTOR, ".ex-contact-box__address-field-fax-number").text if driver.find_elements(By.CSS_SELECTOR, ".ex-contact-box__address-field-fax-number") else ""
            
            # Extract Email and Website
            email = driver.find_element(By.CSS_SELECTOR, ".ex-contact-box__contact-btn").get_attribute("href").split(":")[1].split("?")[0] if driver.find_elements(By.CSS_SELECTOR, ".ex-contact-box__contact-btn") else ""
            website = driver.find_element(By.CSS_SELECTOR, ".ex-contact-box__website-link").get_attribute("href") if driver.find_elements(By.CSS_SELECTOR, ".ex-contact-box__website-link") else ""

            # Storing the data in the dictionary
            data["Company Name"] = company_name
            data["Street Address"] = street_address
            data["City"] = city
            data["ZIP Code"] = zip_code
            data["Country"] = country
            data["Phone"] = telephone
            data["Fax"] = fax
            data["Email"] = email
            data["Website"] = website

            return data  # If data extraction is successful, return it

        except Exception as e:
            print(f"Error while extracting profile data: {e}. Retrying ({attempt + 1}/{retries})...")
            attempt += 1
            time.sleep(2)  # Wait before retrying
    return None  # Return None if retries are exhausted

# Create Excel file and add headers
def create_excel_file():
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Company Data"
    sheet.append(["Company Name", "Street Address", "City", "ZIP Code", "Country", "Phone", "Fax", "Email", "Website"])  # Add headers
    return workbook, sheet

# Save data to Excel
def save_to_excel(workbook, sheet, profile_data):
    if profile_data:  # Only save if profile_data is not None
        sheet.append([ 
            profile_data["Company Name"],
            profile_data["Street Address"],
            profile_data["City"],
            profile_data["ZIP Code"],
            profile_data["Country"],
            profile_data["Phone"],
            profile_data["Fax"],
            profile_data["Email"],
            profile_data["Website"]
        ])
        workbook.save("Company_Data.xlsx")  # Save after each profile is added

# Handle popups and next button navigation
def handle_popup_and_next_button():
    try:
        # Check and close the first popup (if present)
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".js-close-button"))).click()
    except Exception:
        pass  # If first popup is not present, do nothing
    
    try:
        # Check and close the second popup (if present)
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.ex-notification-layer__close"))).click()
    except Exception:
        pass  # If second popup is not present, do nothing

    try:
        # Wait for the next button to be clickable
        next_button = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.btn.btn-default.btn-icon-single span.icon-right")))  # Updated waiting time
        next_button.click()
        time.sleep(3)
    except Exception as e:
        print(f"Error clicking next button: {e}")
        return False  # If next button is not clickable, return False
    return True

# Navigate and scrape company profiles
def scrape_profiles(workbook, sheet):
    current_page = 1  # Start at page 1
    
    # Open the initial page (page 1)
    driver.get(f"https://nordstil.messefrankfurt.com/hamburg/de/ausstellersuche.html?page={current_page}")  # Initial page load
    WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".ex-exhibitor-search-results-container .grid-item")))  # Wait for elements to load
    
    # Handle popup and navigate to the next page (page 2)
    # handle_popup_and_next_button()

    # Move to the third page
    # current_page += 1
    # handle_popup_and_next_button()  # Go to page 2
    # current_page += 1

    # # Move to the fourth page (page 4)
    # current_page += 1
    # handle_popup_and_next_button()  # Go to page 3

    # # Move to the fifth page (page 5)
    # current_page += 1
    # handle_popup_and_next_button()  # Go to page 4

    # # Move to the sixth page (page 6) - Start scraping from here
    # current_page += 1
    # handle_popup_and_next_button()  # Go to page 5
    
    while True:
        # Now scrape from the 6th page
        WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".ex-exhibitor-search-results-container .grid-item")))

        # Scrape the profiles on the current page
        profiles = driver.find_elements(By.CSS_SELECTOR, ".ex-exhibitor-search-results-container .grid-item a")  # Select all company profile links

        for profile in profiles:
            profile_url = profile.get_attribute("href")
            
            # Skip if the profile URL has already been visited
            if profile_url in visited_profiles:
                continue

            visited_profiles.add(profile_url)  # Mark this profile as visited
            driver.execute_script("window.open(arguments[0]);", profile_url)
            driver.switch_to.window(driver.window_handles[1])

            # Extract profile data with retry logic
            profile_data = extract_profile_data()
            if profile_data:  # Only save if profile data was successfully extracted
                print(profile_data)
                save_to_excel(workbook, sheet, profile_data)

            # Close the profile tab
            driver.close()
            driver.switch_to.window(driver.window_handles[0])

        # Handle popups and move to the next page
        if not handle_popup_and_next_button():
            break  # No more pages or an issue with clicking the next button

# Main script
try:
    # Create Excel file and sheet
    workbook, sheet = create_excel_file()
    
    # Scrape profiles and save them to Excel
    scrape_profiles(workbook, sheet)

finally:
    driver.quit()
