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
import time


bind_address = "0.0.0.0"
wait_timeout = 10
poll_frequency = 0.2
tmp_dir = "/tmp"

logging_level = "DEBUG"
logging_format = "%(asctime)s --- %(levelname)s: %(message)s"
url = 'https://zsb.hitwh.edu.cn/home/query/score'

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
    data = {}
    
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'province-box')))
    # 定位省份元素并循环
    province_elements = driver.find_elements(By.XPATH, '/html/body/div[4]/div[2]/div/div[2]/div/div[1]/ul/li')
    province_len = len(province_elements)
    for province_id in range(province_len):
        province_elements = driver.find_elements(By.XPATH, '/html/body/div[4]/div[2]/div/div[2]/div/div[1]/ul/li')
        province_name = province_elements[province_id].text
        print('省份：', province_name)
        province_elements[province_id].click()
    
        time.sleep(2)  # 等待2秒
        
        # 定位年份元素并循环
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'year-box')))
        year_elements = driver.find_elements(By.XPATH, '/html/body/div[4]/div[2]/div/div[2]/div/div[2]/ul/li')
        year_len = len(year_elements)
        for year_id in range(year_len):
            year_elements = driver.find_elements(By.XPATH, '/html/body/div[4]/div[2]/div/div[2]/div/div[2]/ul/li')
            year_name = year_elements[year_id].text
            print("年份名字：",year_name)
            year_elements[year_id].click()
 
            data = get_table_data(driver)
            saveIntoCsv(data,province_name, year_name,driver)

def saveIntoCsv(data,province_name, year_name,driver):
    wb = Workbook()
    # 首先删除默认创建的工作表
    ws = wb.active
    
    # 写入表头数据
    header_data = get_table_header(driver)
    ws.append(header_data)
    
    # 写入表格数据
    for row in data:
        ws.append(row)
    # 保存工作簿到文件
    filename = f"{year_name}-{province_name}-哈尔滨工业大学(威海).csv"
    os.makedirs('学校信息\哈尔滨工业大学(威海)', exist_ok=True)
    wb.save(os.path.join('学校信息\哈尔滨工业大学(威海)', filename))

def get_table_header(driver):
  
    # 找到表头元素
    headers = driver.find_elements(By.XPATH, '/html/body/div[4]/div[2]/div/div[2]/table/thead/tr/td')
    
    # 提取列头文本
    header_data = [column_header.text.strip() for column_header in headers]
    
    print("表格头部信息：", header_data)
    
    return header_data

def get_table_data(driver):
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'tbody')))
    # 找到所有行数据
    rows = driver.find_elements(By.XPATH, '//*[@id="tbody"]/tr')
    print('111111111111',rows)
    
    data = []
    
    # 提取每行数据
    for row in rows:
        # 找到当前行下的所有单元格元素
        cells = row.find_elements(By.XPATH, './td')
        row_data = [cell.text.strip() for cell in cells]
        data.append(row_data)

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
        rpdb.set_trace(bind_address, 0)

if __name__ == "__main__":
    logger = logging.Logger("main", logging_level)
    handler = logging.StreamHandler()
    handler.setFormatter(logging.Formatter(logging_format))
    logger.addHandler(handler)
    run(logger)
