from __future__ import annotations
import copy
from openpyxl import Workbook
import sys, os, time, base64, signal, logging, traceback, rpdb
from argparse import ArgumentParser
from datetime import datetime as dt, timedelta as td
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
import time


bind_address = "0.0.0.0"
wait_timeout = 10
poll_frequency = 0.2
tmp_dir = "/tmp"

logging_level = "DEBUG"
logging_format = "%(asctime)s --- %(levelname)s: %(message)s"
url = 'https://bm.bsu.edu.cn/lscj/index.aspx'

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

    WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, 'kk1')))
    year_elements = driver.find_elements(By.XPATH, '//*[@id="kk1"]/table[2]/tbody/tr/td/div/a')
    year_len = len(year_elements)
    for year_id in range(year_len):
        year_elements = driver.find_elements(By.XPATH, '//*[@id="kk1"]/table[2]/tbody/tr/td/div/a')
        year_name = year_elements[year_id].text
        print('年份：', year_name)
        year_elements[year_id].click()
        data_all = {}
    
        WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, 'kk1')))
        province_elements = driver.find_elements(By.XPATH, '//*[@id="kk1"]/table[3]/tbody/tr/td/div/a')
        province_len = len(province_elements)
        for province_id in range(province_len):
            province_elements = driver.find_elements(By.XPATH, '//*[@id="kk1"]/table[3]/tbody/tr/td/div/a')
            province_name = province_elements[province_id].text
            print('省份：', province_name)
            province_elements[province_id].click()
            
            
            WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, 'kk1')))
            leibie_elements = driver.find_elements(By.XPATH, '//*[@id="kk1"]/table[1]/tbody/tr/td/div/a')
            leibie_len = len(leibie_elements)
            for leibie_id in range(leibie_len):
                leibie_elements = driver.find_elements(By.XPATH, '//*[@id="kk1"]/table[1]/tbody/tr/td/div/a')
                leibie_name = leibie_elements[leibie_id].text
                print('类别：', leibie_name)
                leibie_elements[leibie_id].click()
            
                time.sleep(3)
                data = get_table_data(driver)
                data_all[f'{leibie_name}'] = copy.deepcopy(data)
                save_into_csv(data_all, year_name,province_name,driver)

def save_into_csv(data, year_name, province_name, driver):
    wb = Workbook()
    # 首先删除默认创建的工作表
    default_ws = wb.active
    wb.remove(default_ws)

    for subject_type_key, rows in data.items():
        ws = wb.create_sheet(title=subject_type_key)
        # 添加表头
        combined_header_data = get_table_header(driver)
        ws.append(combined_header_data)
        
        for row in rows:
            ws.append(row)
    
    # 保存文件
    filename = f"{year_name}-{province_name}-北京体育大学.csv"
    os.makedirs('学校信息\北京体育大学', exist_ok=True)
    wb.save(os.path.join('学校信息\北京体育大学', filename))


def get_table_header(driver):
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'table2')))
    header = driver.find_element(By.ID, 'table2')
    column_headers = header.find_elements(By.XPATH, '//*[@id="table2"]/tbody/tr[2]/th[position() !=4]')
    header_data = [column_header.text.strip() for column_header in column_headers]
    print("表格头部信息：", header_data)
    return header_data


def get_table_data(driver):
    try:
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'table2')))
        data_wrapper = driver.find_element(By.ID, 'table2')
    except NoSuchElementException:
        print("未找到数据包装器元素，跳过当前省份处理")
        return []
    rows = data_wrapper.find_elements(By.XPATH, '//*[@id="table2"]/tbody/tr[position() >3]')
    data = []
    for row in rows:
        try:
            cells = row.find_elements(By.XPATH, './td')
            row_data = []
            for cell in cells:
                row_data.append(cell.text.strip())
            data.append(row_data)
        except StaleElementReferenceException:
            rows = data_wrapper.find_elements(By.XPATH, '//*[@id="table2"]/tbody/tr[position() >3]')
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
