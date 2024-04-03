from __future__ import annotations
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
import csv


bind_address = "0.0.0.0"
wait_timeout = 10
poll_frequency = 0.2
tmp_dir = "/tmp"

logging_level = "DEBUG"
logging_format = "%(asctime)s --- %(levelname)s: %(message)s"
url = 'https://zsb.scau.edu.cn/lnlqfs/list.htm'

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
    
    WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.CLASS_NAME, 'year-select')))
    year_elements = driver.find_elements(By.XPATH, '//*[@id="l-container"]/div/div[2]/div/div[2]/div[1]/div[1]/div[2]/select/option[position() = 4]')
    year_len = len(year_elements)
    for year_id in range(year_len):
        year_elements = driver.find_elements(By.XPATH, '//*[@id="l-container"]/div/div[2]/div/div[2]/div[1]/div[1]/div[2]/select/option[position() = 4]')
        year_name = year_elements[year_id].text
        print('年份：', year_name)
        year_elements[year_id].click()
            
        WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.CLASS_NAME, 'city-select')))
        province_elements = driver.find_elements(By.XPATH, '//*[@id="l-container"]/div/div[2]/div/div[2]/div[1]/div[2]/div[2]/select/option[position() > 1]')
        province_len = len(province_elements)
        for province_id in range(province_len):
            province_elements = driver.find_elements(By.XPATH, '//*[@id="l-container"]/div/div[2]/div/div[2]/div[1]/div[2]/div[2]/select/option[position() > 1]')
            province_name = province_elements[province_id].text
            print('省份：', province_name)
             # 如果省份名字是西藏、陕西或宁夏，则跳过该省份
            if province_name in ['西藏', '陕西', '宁夏']:
                continue
            # 循环处理年份
            # 组合省份名字和年份
            combined_name = year_name + ' ' + province_name
                
            # 判断是否需要跳过
            if combined_name in ['2021 内蒙古',"2021 广东"]:
                continue
            province_elements[province_id].click()
    
            clickbutton(driver)
            
            time.sleep(3)
            data = get_table_data(driver)
            save_into_csv(data, year_name,province_name)

def clickbutton(driver):
    button =  WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.XPATH, '//*[@id="l-container"]/div/div[2]/div/div[2]/div[1]/div[3]')))
    button.click()

def save_into_csv(data,year_name,province_name):

    wb = Workbook()
    # 首先删除默认创建的工作表
    default_ws = wb.active
    wb.remove(default_ws)

    # 创建工作表并写入数据
    ws = wb.create_sheet(title="Data")
    
    # 写入数据
    for row in data:
        ws.append(row)
    
    # 保存文件
    filename = f"{year_name}-{province_name}-华南农业大学.csv"
    os.makedirs('学校信息\华南农业大学', exist_ok=True)
    wb.save(os.path.join('学校信息\华南农业大学', filename))

def get_table_data(driver):
    WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.CLASS_NAME, 'search-text')))
    data_wrapper = driver.find_element(By.CLASS_NAME, 'search-text')

    rows = data_wrapper.find_elements(By.XPATH, '//*[@id="l-container"]/div/div[2]/div/div[2]/div[2]/div[1]/div[2]/table/tbody/tr')
    data = []
    for row in rows:
        try:
            cells = row.find_elements(By.XPATH, './td')
            row_data = []
            for cell in cells:
                row_data.append(cell.text.strip())
            data.append(row_data)
        except StaleElementReferenceException:
            rows = data_wrapper.find_elements(By.XPATH, '//*[@id="l-container"]/div/div[2]/div/div[2]/div[2]/div[1]/div[2]/table/tbody/tr')
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
