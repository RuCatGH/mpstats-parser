import pickle
import os
import time
from typing import List

from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium import webdriver
from dotenv import load_dotenv
import pandas as pd
from dotenv import load_dotenv

load_dotenv()

login = os.getenv('LOGIN')
password = os.getenv('PASSWORD')

# Функция для инициализации драйвера Chrome с настройками
def initialize_chrome_driver() -> webdriver.Chrome:

    current_directory = os.getcwd()
    options = Options()
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_argument("--start-maximized")
    options.add_argument("--profile.default_content_settings.popups=0")
    options.add_argument("--profile.default_content_setting_values.automatic_downloads=1")
    options.add_argument("--disable-blink-features")
    preferences = {"download.default_directory": current_directory,
                   "download.prompt_for_download": False,
                   "directory_upgrade": True,
                   "safebrowsing.enabled": True }
    options.add_experimental_option("prefs", preferences)

    options.add_argument(f"--download.default_directory={current_directory}")
    service = Service(executable_path=current_directory + '\chromedriver.exe')

    driver = webdriver.Chrome(options=options, service=service)
    return driver

# Функция для получения ссылок на товары с сайта Ozon
def get_ozon_product_links(driver: webdriver.Chrome, queries: List[str]) -> List[str]:
    all_links = []
    driver.get('https://www.ozon.ru/')
    time.sleep(2)
    for query in queries:
        driver.get(f"https://www.ozon.ru/search/?from_global=true&text={query.replace(' ', '+')}")
        time.sleep(1)
        links = set([item.get_attribute('href').split('/')[4].split('-')[-1] for item in driver.find_elements(By.CLASS_NAME, 'tile-hover-target')[:60]])
        all_links.extend(links)
    return all_links

# Функция для авторизации на сайте mpstats.club и сохранения куков
def login_and_save_cookies(driver: webdriver.Chrome, login: str, password: str):
    driver.get('https://mpstats.club/login')
    if not os.path.exists('cookies'):
        wait = WebDriverWait(driver, 10)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".d-none.d-xxl-block.btn"))).click()
        wait.until(EC.presence_of_element_located((By.NAME, "username"))).send_keys(login)
        wait.until(EC.presence_of_element_located((By.NAME, "password"))).send_keys(password)
        time.sleep(1)
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, "btn"))).click()
        time.sleep(90)
        pickle.dump(driver.get_cookies(), open('cookies', 'wb'))

# Функция для добавления куков к драйверу
def add_cookies_to_driver(driver: webdriver.Chrome):
    for cookie in pickle.load(open('cookies', 'rb')):
        driver.add_cookie(cookie)


# Функция для выполнения запросов на mpstats.club и сохранения результатов в Excel
def perform_mpstats_requests(driver: webdriver.Chrome, all_links: List[str], workbook_pd):
    current_directory = os.getcwd()
    wait = WebDriverWait(driver, 30)
    queries = {}
    for batch in zip(range(0, len(all_links), 30), workbook_pd['Наименование'].to_list()):
        try:
            driver.get('https://mpstats.club/seo/keywords/expanding')
            wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(., 'Ozon')]"))).click()
            wait.until(EC.presence_of_element_located((By.CLASS_NAME, "form-control"))).send_keys('\n'.join(all_links[batch[0]:batch[0]+30]))
            wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Получить отчет']"))).click()

            export = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "btn-xs")))
            driver.execute_script("arguments[0].scrollIntoView();", export)
            export.click()
            time.sleep(1)
            try:
                download_button = driver.find_element(By.XPATH, "//header[text()='Xlsx формат:']/following-sibling::ul/li/button").click()
                time.sleep(1)

                for file in os.listdir(current_directory):
                    if file.startswith("SEO") and file.endswith(".xlsx"):
                        file_path = os.path.join(current_directory, file)
                        
                        # Откройте текущий файл
                        workbook = pd.read_excel(file)
                        
                        queries[batch[1]] ={
                            'Запросы': workbook['Запросы'].tolist()[:-1],
                            'Количество запросов на Ozon': workbook['Частота Oz'].tolist()[:-1],
                            'Количество запросов на WB': workbook['Частота WB'].tolist()[:-1]
                            }

            
                        os.remove(file_path)
            except Exception as ex:
                time.sleep(10)
                continue
        except:
            continue
    expanded_data = []
    for _, row in workbook_pd.iterrows():
        name_query = queries.get(row['Наименование'], {})
        try:
            for i in range(len(name_query['Запросы'])):
                запрос = name_query['Запросы'][i]
                query_ozon = name_query['Количество запросов на Ozon'][i]
                query_wb = name_query['Количество запросов на WB'][i]
                
                expanded_data.append([row['Наименование'], row['Ключ'], запрос, query_ozon, query_wb])
        except:
            continue

    # Создаем новый DataFrame с дополнительными полями
    expanded_df = pd.DataFrame(expanded_data, columns=['Наименование', 'Ключ', 'Запрос', 'Количество запросов на Ozon', 'Количество запросов на WB'])

    expanded_df.to_excel('expanded_data.xlsx', index=False)

# Главная функция для выполнения всех шагов
def main():
    try:
        df = pd.read_excel('Ключи.xlsx')
        queries = df['Ключ'].tolist()

        driver = initialize_chrome_driver()
        login_and_save_cookies(driver, login, password)
        add_cookies_to_driver(driver)

        all_links = get_ozon_product_links(driver, queries)
        perform_mpstats_requests(driver, all_links, df)
    except Exception as ex:
        print(ex)
        time.sleep(50)


if __name__ == "__main__":
    main()
