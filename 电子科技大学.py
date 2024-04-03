from __future__ import annotations
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
from openpyxl import Workbook
import copy
import time

bind_address = "0.0.0.0"
wait_timeout = 10
poll_frequency = 0.2
tmp_dir = "/tmp"

logging_level = "DEBUG"
logging_format = "%(asctime)s --- %(levelname)s: %(message)s"
url = 'https://www.zs.uestc.edu.cn/benchmark/'

os.environ["CUDA_VISIBLE_DEVICES"] = "-1"
os.environ["TF_CPP_MIN_LOG_LEVEL"] = "3"
for proxy in ["https_proxy", "http_proxy", "socks_proxy", "all_proxy"]:
    os.environ[proxy] = os.environ[proxy.upper()] = ""

_driver: WebDriver
_logger: logging.Logger

def restart_session(display) -> None:
    global _driver
    _driver = new_session(display)

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
    all_data = []
    time.sleep(5)
    
    WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.CLASS_NAME, 'formitem')))
    year_elements = driver.find_elements(By.XPATH, '//*[@id="form1"]/div[3]/select/option[text() = "2023年" or text() = "2022年" or text() ="2021年"]')
    year_len = len(year_elements)
    for year_id in range(year_len):
        year_elements = driver.find_elements(By.XPATH, '//*[@id="form1"]/div[3]/select/option[text() = "2023年" or text() = "2022年" or text() ="2021年"]')
        year_name = year_elements[year_id].text
        print('年份：', year_name)
        year_elements[year_id].click()
    
        time.sleep(3)
        data = get_table_data(driver)
        if data:
            saveIntoXlsx(data, year_name,driver)
        if not is_last_page(driver):
            button(driver)
            data1 = get_table_data(driver)
            all_data.extend(data1)
    
        saveIntoXlsx(all_data, year_name,driver)

def button(driver):
    next_button = driver.find_element(By.XPATH, '//*[@id="Pagination"]/span[2]')
    next_button.click()
    
def is_last_page(driver):
    try:
        next_button = driver.find_element(By.XPATH, '//*[@id="Pagination"]/span[2]')
        return False
    except:
        return True

def saveIntoXlsx(data, year_name,driver):
    wb = Workbook()
    # 首先删除默认创建的工作表
    default_ws = wb.active
    wb.remove(default_ws)

    # 创建工作表并写入数据
    ws = wb.create_sheet(title="Data")
    
    header_data = get_table_header(driver)
    ws.append(header_data)
    # 写入数据
    for row in data:
        ws.append(row)

    # 保存工作簿到文件
    filename = f"{year_name}-甘肃-电子科技大学.csv"
    os.makedirs('学校信息\电子科技大学', exist_ok=True)
    wb.save(os.path.join('学校信息\电子科技大学', filename))
    
def get_table_header(driver):
    # 等待表格数据包装器元素出现
    WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, 'tpl_holder')))  
    
    # 找到表头元素
    table = driver.find_element(By.ID, 'tpl_holder')
    headers = table.find_elements(By.XPATH, '//*[@id="tpl_holder"]/table/thead/tr/th')
    
    # 提取列头文本
    header_data = [column_header.text.strip() for column_header in headers]
    
    print("表格头部信息：", header_data)
    
    return header_data

def get_table_data(driver):

    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'tpl_holder')))
    data_wrapper = driver.find_element(By.ID, 'tpl_holder')

    rows = data_wrapper.find_elements(By.XPATH, '//*[@id="tpl_holder"]/table/tbody/tr')
    data = []
    for row in rows:
        try:
            cells = row.find_elements(By.XPATH, './td')
            row_data = []
            for cell in cells:
                row_data.append(cell.text.strip())
            data.append(row_data)
        except StaleElementReferenceException:
            rows = data_wrapper.find_elements(By.XPATH, '//*[@id="tpl_holder"]/table/tbody/tr')
            continue
    print("表格数据：", data)
    return data

def run(logger: logging.Logger):
    global _logger
    _logger = logger
    try:
        restart_session('headless')
        provinces = get_provinces(_driver)
        logger.info("success")
    except KeyboardInterrupt:
        os.abort()
    except RuntimeError:
        rpdb.set_trace(bind_address, 0)
        return
    except:
        logger.critical(traceback.format_exc())
        # logger.info(str(locals()))
        rpdb.set_trace(bind_address, 0)

if __name__ == "__main__":
    logger = logging.Logger("main", logging_level)
    handler = logging.StreamHandler()
    handler.setFormatter(logging.Formatter(logging_format))
    logger.addHandler(handler)
    run(logger)
