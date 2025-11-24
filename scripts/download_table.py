import os
import time
import logging
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# set up logging config
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Environment variables
username = os.getenv("username")
password = os.getenv("password")

if not username or not password:
    raise ValueError("Environment variables 'USER_NAME' and/or 'USER_PASSWORD' not set.")

# --- DATE LOGIC: get from 1st of current month until today ---
today = datetime.now()
first_day_of_month = today.replace(day=1)

inicio = first_day_of_month.strftime("%d%m%Y")  # always 1st day of current month
fim = today.strftime("%d%m%Y")  # current day

logging.info(f"Using date range: {inicio} to {fim}")

# --- DOWNLOAD DIRECTORY ---
download_dir = os.getcwd()

# --- CHROME OPTIONS ---
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("--start-maximized")

prefs = {
    "download.default_directory": download_dir,
    "download.prompt_for_download": False,
    "directory_upgrade": True,
    "safebrowsing.enabled": False,
    "safebrowsing.disable_download_protection": True
}
chrome_options.add_experimental_option("prefs", prefs)
chrome_options.add_argument("--unsafely-treat-insecure-origin-as-secure=http://drogcidade.ddns.net:4647/sgfpod1/Login.pod")

# --- START SELENIUM ---
driver = webdriver.Chrome(options=chrome_options)

try:
    logging.info("Navigating to target URL and logging in...")
    driver.get("http://drogcidade.ddns.net:4647/sgfpod1/Login.pod")

    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "id_cod_usuario"))).send_keys(username)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "nom_senha"))).send_keys(password)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "login"))).click()

    WebDriverWait(driver, 10).until(lambda x: x.execute_script("return document.readyState === 'complete'"))
    time.sleep(5)

    # verifica se o pop-up esta presente e clica se ele existir
    """popup_element = driver.find_element(By.ID, "modalMsgMovimentacaoAnvisa") 
    if popup_element:
        driver.find_element(By.ID, "sairModalMsgMovimentos").click()"""

    # Access "Vendas por Data"
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "sideMenuSearch")))
    driver.find_element(By.ID, "sideMenuSearch").send_keys("Vendas por Data")
    driver.find_element(By.ID, "sideMenuSearch").click()
    driver.implicitly_wait(2)

    driver.find_element(By.CSS_SELECTOR, '[title="Vendas por Data"]').click()
    WebDriverWait(driver, 10).until(lambda x: x.execute_script("return document.readyState === 'complete'"))

    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "agrup_fil_2"))).click()
    time.sleep(5)
    ts = time.strftime("%Y%m%d-%H%M%S")
    driver.save_screenshot(f"screenshot_{ts}.png")

    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "tabTabdhtmlgoodies_tabView1_1"))).click()
    time.sleep(5)

    # Fill in start and end dates
    driver.find_element(By.ID, "dat_inicio").send_keys(inicio)
    driver.find_element(By.ID, "dat_fim").send_keys(fim)
    time.sleep(5)

    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "selI_1"))).click()
    time.sleep(5)

    # Report format (XLS)
    driver.find_element(By.ID, "saida_4").click()

    # Trigger download
    logging.info("Triggering report download...")
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "runReport"))).click()
    logging.info("Download has started.")

    # Wait for download to complete
    time.sleep(15)

    # Get the most recent downloaded file
    files = os.listdir(download_dir)
    downloaded_files = [f for f in files if f.endswith('.xls')]
    if downloaded_files:
        downloaded_files.sort(key=lambda x: os.path.getmtime(os.path.join(download_dir, x)))
        most_recent_file = downloaded_files[-1]
        downloaded_file_path = os.path.join(download_dir, most_recent_file)

        file_size = os.path.getsize(downloaded_file_path)
        logging.info(f"Download completed successfully. File path: {downloaded_file_path}, Size: {file_size} bytes")
    else:
        logging.error("Download failed. No .xls files found.")

finally:
    driver.quit()
