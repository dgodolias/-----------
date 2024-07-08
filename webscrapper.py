from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd

def scrape_doctor_info(url):
    """Scrapes doctor information from a given URL using Selenium."""

    driver = webdriver.Chrome()  # Use the appropriate browser driver
    driver.get(url)

    # Wait for the page to load (optional)
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'CompanyNameLbl')))
    except:
        print(f"Error loading page: {url}")
        driver.quit()
        return None, None, None, None, None, None, None

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

    # Extract website
    website_element = soup.find('a', class_='rc_Detaillink', href=True, itemprop='url')
    website = website_element['href'] if website_element else None

    # Extract email
    email_element = soup.find('a', rel='nofollow', class_='rc_Detaillink', href=True)
    email = email_element['href'].replace('mailTo:', '') if email_element else None

    driver.quit()  # Close the browser

    return name, address, profession, phone, mobile, website, email

def main():
    """Reads URLs from a file, scrapes doctor information, and writes it to an Excel file."""

    with open('urls.txt', 'r') as f:
        urls = [line.strip() for line in f]

    doctor_data = []
    for url in urls:
        # Skip blank lines
        if not url:
            continue

        name, address, profession, phone, mobile, website, email = scrape_doctor_info(url)
        if name:
            doctor_data.append([name, address, profession, phone, mobile, website, email])

    df = pd.DataFrame(doctor_data, columns=['Name', 'Address', 'Profession', 'Phone', 'Mobile', 'Website', 'Email'])
    df.to_excel('doctor_info.xlsx', index=False)

if __name__ == '__main__':
    main()
