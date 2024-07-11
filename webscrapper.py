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

def get_profile_links(search_url):
    """Fetches profile links from the search results page."""
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

    chrome_driver_path = './chromedriver-win64/chromedriver.exe'

    driver = webdriver.Chrome(service=Service(chrome_driver_path), options=chrome_options)
    driver.get(search_url)

    try:
        WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'AdvAreaRight')))
    except Exception as e:
        print(f"Error loading search results page: {search_url}\n{e}")
        driver.quit()  # Close the driver if an error occurs
        return []

    time.sleep(2)

    soup = BeautifulSoup(driver.page_source, 'html.parser')
    profile_links = []

    for div in soup.find_all('div', class_='AdvAreaRight'):
        a_tag = div.find('a', class_='AdvAreaBottomRight')
        if a_tag and 'href' in a_tag.attrs:
            profile_links.append(a_tag['href'])

    driver.quit()  # Close the driver after fetching links
    return profile_links

def scrape_doctor_info(url):
    """Scrapes doctor information from a given URL using Selenium."""

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

    chrome_driver_path = './chromedriver-win64/chromedriver.exe'

    driver = webdriver.Chrome(service=Service(chrome_driver_path), options=chrome_options)
    driver.get(url)

    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'CompanyNameLbl')))
    except Exception as e:
        print(f"Error loading page: {url}\n{e}")
        driver.quit()  # Close the driver if an error occurs
        return None, None, None, None, None, None, None, None

    time.sleep(2)

    soup = BeautifulSoup(driver.page_source, 'html.parser')

    name_element = soup.find('label', id='CompanyNameLbl', class_='companyLabel_class', itemprop='name')
    name = name_element.span.text.strip() if name_element else None

    address_element = soup.find('label', id='AddressLbl')
    address = address_element.text.strip() if address_element else None

    profession_element = soup.find('label', id='ProfessionLbl', itemprop='description')
    profession = profession_element.text.strip() if profession_element else None

    phone_element = soup.find('label', class_='rc_firstphone')
    phone = phone_element.text.strip() if phone_element else None

    mobile_element = soup.find('label', id='MobileContLbl')
    mobile = mobile_element.span.text.strip() if mobile_element else None

    website_element = soup.find('a', class_='rc_Detaillink', href=True, itemprop='url')
    website = website_element['href'] if website_element else None

    email_element = soup.find('a', rel='nofollow', class_='rc_Detaillink', href=True)
    email = email_element['href'].replace('mailto:', '') if email_element else None

    # Add a blank value for the new column
    hours = ''

    driver.quit()  # Close the driver after scraping doctor info
    return name, address, profession, phone, mobile, website, email, hours

def scrape_and_append(url, df):
    """Scrapes doctor information and appends it to the doctor_data list if not a duplicate."""
    global doctor_data
    name, address, profession, phone, mobile, website, email, hours = scrape_doctor_info(url)

    if name and not check_phone_exists(df, phone, mobile):
        doctor_data.append([name, address, profession, phone, mobile, website, email, hours])

    driver.quit()  # Close the driver after scraping each doctor's info

def check_phone_exists(df, phone, mobile):
    """Check if the phone or mobile number already exists in the DataFrame."""
    phone = str(phone)
    mobile = str(mobile)
    return ((df["Phone"].astype(str) == phone) | (df["Mobile"].astype(str) == mobile)).any()

def main():
    """Reads URLs from a file, scrapes doctor information, and writes it to an Excel file."""

    global doctor_data
    doctor_data = []

    doctor_type = input("Enter the type of doctors (in Greek): ")
    location = input("Enter the location (in Greek): ")
    search_url = f"https://www.vrisko.gr/search/{doctor_type}/{location}"

    profile_links = get_profile_links(search_url)
    if not profile_links:
        print("No profile links found.")
        return

    try:
        df = pd.read_excel('doctor_info.xlsx')
        doctor_data = df.values.tolist()
    except FileNotFoundError:
        df = pd.DataFrame(columns=['Name', 'Address', 'Profession', 'Phone', 'Mobile', 'Website', 'Email', 'Ωρες'])

    max_threads = 8
    active_threads = []

    for url in profile_links:
        if not url:
            continue

        while len(active_threads) >= max_threads:
            for thread in active_threads:
                if not thread.is_alive():
                    active_threads.remove(thread)
            time.sleep(0.1)

        thread = threading.Thread(target=scrape_and_append, args=(url, df))
        active_threads.append(thread)
        thread.start()

    for thread in active_threads:
        thread.join()

    updated_df = pd.DataFrame(doctor_data, columns=['Name', 'Address', 'Profession', 'Phone', 'Mobile', 'Website', 'Email', 'Ωρες'])
    df = pd.concat([df, updated_df]).drop_duplicates(subset=['Phone', 'Mobile'], keep='first')

    df.to_excel('doctor_info.xlsx', index=False)
    print("Data appended to doctor_info.xlsx")

if __name__ == '__main__':
    main()
