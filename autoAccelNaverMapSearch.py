import time
import re
import xlwings as xw
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

def switch_frame(driver, frame):
    driver.switch_to.default_content() 
    driver.switch_to.frame(frame)

wb = xw.Book('testaccel.xlsx')
sheet = wb.sheets['Sheet1']  

chrome_driver_path = 'chromedriver-win64\\chromedriver.exe' 

chrome_options = Options()
chrome_options.add_argument('--disable-gpu')  
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')
chrome_options.add_argument('--window-size=1920x1080') 
service = Service(chrome_driver_path)

driver = webdriver.Chrome(service=service, options=chrome_options)

for idx, address in enumerate(sheet.range('D2:D{}'.format(sheet.range('D1').end('down').row)).value, start=2):
    if not address:
        continue  
    
    full_address = f"{address}"
    search_url = f"https://map.naver.com/v5/search/{full_address}"
    driver.get(search_url)
    
    try:
        iframe_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'iframe[src*="map.naver.com"]')))
        
        switch_frame(driver, iframe_element)
        
        result_elements = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.YwYLL, span.GHAhO')))
        first_address = re.sub(r'\bsearhString\b', '', result_elements[0].text.strip())
        print(first_address)
        sheet.range('O{}'.format(idx)).value = first_address
    except Exception as e:
        print(f"주소: {full_address}, 오류: {e}")
        sheet.range('O{}'.format(idx)).value = "NG"  


driver.quit()

wb.save('updated_companies.xlsx')
wb.close()
