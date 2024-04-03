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

bind_address = "0.0.0.0"
wait_timeout = 10
poll_frequency = 0.2
tmp_dir = "/tmp"

logging_level = "DEBUG"
logging_format = "%(asctime)s --- %(levelname)s: %(message)s"
url = 'https://www.zs.cdut.edu.cn/lnfs.jsp?urltype=tree.TreeTempUrl&wbtreeid=1038'

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
    
    WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, 'province')))
    province_elements = driver.find_elements(By.XPATH, '//*[@id="province"]/a')
    province_len = len(province_elements)
    print(province_len)
    for province_id in range(province_len):
        province_elements = driver.find_elements(By.XPATH, '//*[@id="province"]/a')
        province_name = province_elements[province_id].text
        print('省份：', province_name)
        province_elements[province_id].click()
        
        WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, 'zsnf')))
        year_elements = driver.find_elements(By.XPATH, '//*[@id="zsnf"]/a')
        year_len = len(year_elements)
        for year_id in range(year_len):
            year_elements = driver.find_elements(By.XPATH, '//*[@id="zsnf"]/a')
            year_name = year_elements[year_id].text
            print('年份：', year_name)
            year_elements[year_id].click()
            
            WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, 'klmc')))
            subject_elements = driver.find_elements(By.XPATH, '//*[@id="klmc"]/a')
            subject_len = len(subject_elements)
            
            data_all = {}
            
            for subject_id in range(subject_len):
                subject_elements = driver.find_elements(By.XPATH, '//*[@id="klmc"]/a')
                subject_name = subject_elements[subject_id].text
                print('科类：', subject_name)
                subject_elements[subject_id].click()
                
                WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, 'zslx')))
                type_elements = driver.find_elements(By.XPATH, '//*[@id="zslx"]/a')
                type_len = len(type_elements)
                for type_id in range(type_len):
                    type_elements = driver.find_elements(By.XPATH, '//*[@id="zslx"]/a')
                    type_name = type_elements[type_id].text
                    print('类型：', type_name)
                    type_elements[type_id].click()
                    data = get_table_data(driver)
                    data_all[f'{subject_name}_{type_name}'] = copy.deepcopy(data)
            saveIntoXlsx(data_all, province_name, year_name,driver)

def saveIntoXlsx(data, province_name, year_name,driver):
    wb = Workbook()
    # 首先删除默认创建的工作表
    default_ws = wb.active
    wb.remove(default_ws)
    
    for subject_type_key, rows in data.items():
        ws = wb.create_sheet(title=clean_sheet_title(subject_type_key))
        header_data = get_table_header(driver)
        ws.append(header_data)
        for row in rows:
            ws.append(row)

    # 保存工作簿到文件
    filename = f"{year_name}-{province_name}-成都理工大学.csv"
    os.makedirs('学校信息\成都理工大学', exist_ok=True)
    wb.save(os.path.join('学校信息\成都理工大学', filename))

def clean_sheet_title(title):
    # 使用 replace 方法过滤掉斜杠
    return title.replace('/', '')

def get_table_header(driver):
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'sszygradeListPlace')))
    header = driver.find_element(By.ID, 'sszygradeListPlace')
    column_headers = header.find_elements(By.XPATH, '//*[@id="sszygradeListPlace"]/thead/tr/th')
    header_data = [column_header.text.strip() for column_header in column_headers]
    print("表格头部信息：", header_data)
    return header_data

def get_table_data(driver):
    WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, 'zsjhfzyqk')))
    table = driver.find_element(By.ID, 'zsjhfzyqk')
    # 使用XPath选择所有td和th元素
    rows = table.find_elements(By.XPATH, './/tr')
    data = []
    for row in rows:
        try:
            cells = row.find_elements(By.XPATH, './/td')
            row_data = []
            for cell in cells:
                row_data.append(cell.text.strip())
            data.append(row_data)
        except StaleElementReferenceException:
            rows = table.find_elements(By.ID, 'zsjhfzyqk')
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
