from __future__ import annotations
import copy
from openpyxl import Workbook
import os, time, logging, traceback, rpdb
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import NoSuchElementException
import time


bind_address = "0.0.0.0"
wait_timeout = 10
poll_frequency = 0.2
tmp_dir = "/tmp"

logging_level = "DEBUG"
logging_format = "%(asctime)s --- %(levelname)s: %(message)s"
url = 'https://zhinengdayi.com/page/detail/NJBCYO/4894/13601'

os.environ["CUDA_VISIBLE_DEVICES"] = "-1"
os.environ["TF_CPP_MIN_LOG_LEVEL"] = "3"
for proxy in ["https_proxy", "http_proxy", "socks_proxy", "all_proxy"]:
    os.environ[proxy] = os.environ[proxy.upper()] = ""

_driver: WebDriver
_logger: logging.Logger

def restart_session(display) -> None:
    global _driver
    _driver = new_session(display)
    return _driver

def wait_find(driver: WebDriver, by: str, value: str, all=True):
    if all:
        finder = lambda d: d.find_elements(by, value)
    else:
        finder = lambda d: d.find_element(by, value)
    return WebDriverWait(driver, wait_timeout, poll_frequency).until(finder)

def new_session(display) -> WebDriver:
    options = FirefoxOptions()
    if display == 'headless':
        options.headless = True
    # 设置 Firefox 可执行文件的路径
    firefox_binary = "C:/Program Files/Mozilla Firefox/firefox.exe"
    options.binary_location = firefox_binary
    driver = webdriver.Firefox(options=options)
    return driver

def get_provinces(driver):
    driver.get(url)
     
    WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, 'cityName')))
    province_elements = driver.find_elements(By.XPATH, '//*[@id="cityName"]/option')
    province_len = len(province_elements)
    for province_id in range(province_len):
        province_elements = driver.find_elements(By.XPATH, '//*[@id="cityName"]/option')
        province_name = province_elements[province_id].text
        print('生源地：', province_name)
        province_elements[province_id].click()

        WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, 'scienceClass')))
        kelei_elements = driver.find_elements(By.XPATH, '//*[@id="scienceClass"]/option')
        kelei_elements_len = len(kelei_elements)
        for kelei_id in range(kelei_elements_len):
            kelei_elements = driver.find_elements(By.XPATH, '//*[@id="scienceClass"]/option')
            kelei_name = kelei_elements[kelei_id].text
            print('科类：', kelei_name)
            kelei_elements[kelei_id].click()

            time.sleep(3)
            data = get_table_data(driver)
            save_into_csv(data,province_name,driver)

def save_into_csv(data, province_name, driver):
    wb = Workbook()
    # 首先删除默认创建的工作表
    ws = wb.active
    
    # 写入表头数据
    header_data = get_table_header(driver)
    ws.append(header_data)
    
    # 写入表格数据
    for row in data:
        ws.append(row)
    
    filename = f"{province_name}-湖南师范大学.csv"
    os.makedirs('学校信息\湖南师范大学', exist_ok=True)
    wb.save(os.path.join('学校信息\湖南师范大学', filename))

def get_table_header(driver):
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'schoolParamList')))
    header = driver.find_element(By.ID, 'schoolParamList')
    column_headers = header.find_elements(By.XPATH, '//*[@id="schoolParamList"]/tbody/tr/th')
    header_data = [column_header.text.strip() for column_header in column_headers]
    print("表格头部信息：", header_data)
    return header_data


def get_table_data(driver):
    try:
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'schoolParamList')))
        data_wrapper = driver.find_element(By.ID, 'schoolParamList')
    except NoSuchElementException:
        print("未找到数据包装器元素，跳过当前省份处理")
        return []
    rows = data_wrapper.find_elements(By.XPATH, '//*[@id="schoolParamList"]/tbody/tr[position() > 1]')
    data = []
    for row in rows:
        try:
            cells = row.find_elements(By.XPATH, './td')
            row_data = []
            for cell in cells:
                row_data.append(cell.text.strip())
            data.append(row_data)
        except StaleElementReferenceException:
            rows = data_wrapper.find_elements(By.XPATH, '//*[@id="schoolParamList"]/tbody/tr')
            continue
    print("表格数据：", data)
    return data

def run(logger: logging.Logger):
    global _logger
    _logger = logger
    try:
        _driver = restart_session('headless')  # 修改此行
        get_provinces(_driver)
        logger.info("success")
    except KeyboardInterrupt:
        os.abort()
    except RuntimeError:
        rpdb.set_trace(bind_address, 0)
        return
    except:
        logger.critical(traceback.format_exc())

if __name__ == "__main__":
    logger = logging.Logger("main", logging_level)
    handler = logging.StreamHandler()
    handler.setFormatter(logging.Formatter(logging_format))
    logger.addHandler(handler)
    run(logger)
