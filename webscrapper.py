from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd
import threading
import time
import psutil

def init_driver():
    """Initialize the Chrome driver with headless options."""
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("start-maximized")
    chrome_options.add_argument("disable-infobars")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("window-size=1920,1080")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")

    return webdriver.Chrome(options=chrome_options)

def get_current_chrome_processes():
    """Get the list of current Chrome process PIDs."""
    chrome_pids = []
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] in ('chrome.exe', 'chromedriver.exe'):
            chrome_pids.append(proc.info['pid'])
    return set(chrome_pids)

def kill_new_chrome_processes(initial_pids):
    """Kill Chrome processes that were started after the initial process scan."""
    current_pids = get_current_chrome_processes()
    new_pids = current_pids - initial_pids
    for pid in new_pids:
        try:
            proc = psutil.Process(pid)
            proc.kill()
        except psutil.NoSuchProcess:
            pass

def scrape_doctor_info(url):
    """Scrapes doctor information from a given URL using Selenium."""
    driver = init_driver()
    try:
        driver.get(url)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'CompanyNameLbl')))
        time.sleep(2)  # Add a delay to mimic real user behavior

        soup = BeautifulSoup(driver.page_source, 'html.parser')

        # Extract doctor's name
        name_element = soup.find('label', id='CompanyNameLbl', class_='companyLabel_class', itemprop='name')
        name = name_element.span.text.strip() if name_element else None

        # Extract address
        address_element = soup.find('label', id='AddressLbl')
        address = address_element.text.strip() if address_element else None

        # Extract profession
        profession_element = soup.find('label', id='ProfessionLbl', itemprop='description')
        profession = profession_element.text.strip() if profession_element else None

        # Extract phone number
        phone_element = soup.find('label', class_='rc_firstphone')
        phone = phone_element.text.strip() if phone_element else None

        # Extract mobile number
        mobile_element = soup.find('label', id='MobileContLbl')
        mobile = mobile_element.span.text.strip() if mobile_element else None

        # Extract email
        email_element = soup.find('a', rel='nofollow', class_='rc_Detaillink', href=True)
        email = email_element['href'].replace('mailto:', '') if email_element else None

        return name, address, profession, phone, mobile, email
    except Exception as e:
        print(f"Error loading page: {url}\n{e}")
        return None, None, None, None, None, None
    finally:
        driver.quit()
        time.sleep(2)  # Ensure driver.quit() has time to close the browser

def scrape_and_append(url, df, lock):
    """Scrapes doctor information and appends it to the doctor_data list if not a duplicate."""
    name, address, profession, phone, mobile, email = scrape_doctor_info(url)
    if name and not check_phone_exists(df, phone, mobile):
        with lock:
            doctor_data.append([name, address, profession, phone, mobile, email, ""])  # Add blank Ωρα column

def check_phone_exists(df, phone, mobile):
    """Check if the phone or mobile number already exists in the DataFrame."""
    phone = str(phone)
    mobile = str(mobile)
    return ((df["Phone"].astype(str) == phone) | (df["Mobile"].astype(str) == mobile)).any()

def get_doctor_links(driver, search_type, search_location):
    """Get doctor links from vrisko.gr based on search type and location."""
    url = f"https://www.vrisko.gr/search/{search_type}/{search_location}"
    driver.get(url)

    soup = BeautifulSoup(driver.page_source, 'html.parser')
    doctor_links = []
    
    for div in soup.find_all('div', class_='AdvAreaRight'):
        website_link = div.find('a', class_='urlClickLoggingClass', target='_blank')
        if not website_link:
            more_link = div.find('a', class_='AdvAreaBottomRight')
            if more_link:
                doctor_links.append(more_link['href'])

    return doctor_links

def main():
    """Reads profession and area, scrapes doctor information, and writes it to an Excel file."""
    global doctor_data
    doctor_data = []
    lock = threading.Lock()

    initial_chrome_pids = get_current_chrome_processes()

    try:
        # Attempt to read existing data from the Excel file
        df = pd.read_excel('doctor_info.xlsx')
        doctor_data = df.values.tolist()
    except FileNotFoundError:
        # File not found, create an empty DataFrame with the expected columns
        df = pd.DataFrame(columns=['Name', 'Address', 'Profession', 'Phone', 'Mobile', 'Email', 'Ωρα'])

    # Initialize the web driver
    driver = init_driver()

    # Get user input for search type and location
    search_type = input("Enter the type of doctor (in Greek, e.g., Ορθοδοντικοί): ")
    search_location = input("Enter the location (in Greek, e.g., Αττική): ")

    # Get the list of doctor links
    doctor_links = get_doctor_links(driver, search_type, search_location)

    # Quit the initial driver used to get links
    driver.quit()
    time.sleep(2)  # Ensure driver.quit() has time to close the browser

    # Limit the number of threads to 8
    max_threads = 8
    active_threads = []

    for url in doctor_links:
        # Skip blank lines
        if not url:
            continue

        # Wait for an available thread slot
        while len(active_threads) >= max_threads:
            for thread in active_threads:
                if not thread.is_alive():
                    active_threads.remove(thread)
            time.sleep(0.1)

        thread = threading.Thread(target=scrape_and_append, args=(url, df, lock))
        active_threads.append(thread)
        thread.start()

    # Wait for all threads to finish
    for thread in active_threads:
        thread.join()

    # Convert the updated list back to a DataFrame
    updated_df = pd.DataFrame(doctor_data, columns=['Name', 'Address', 'Profession', 'Phone', 'Mobile', 'Email', 'Ωρα'])

    # Append new data to the existing DataFrame and drop duplicates
    df = pd.concat([df, updated_df]).drop_duplicates(subset=['Phone', 'Mobile'], keep='first')

    df.to_excel('doctor_info.xlsx', index=False)
    print("Data appended to doctor_info.xlsx")

    kill_new_chrome_processes(initial_chrome_pids)  # Ensure only new processes are killed

if __name__ == '__main__':
    main()
