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
import csv
import pandas as pd

bind_address = "0.0.0.0"
wait_timeout = 10
poll_frequency = 0.2
tmp_dir = "/tmp"

logging_level = "DEBUG"
logging_format = "%(asctime)s --- %(levelname)s: %(message)s"
url = 'http://bkzs.hfut.edu.cn/static/front/hfut/basic/html_web/lnfs.html?param=%u5B89%u5FBD_%7B%7BnowYear%7D%7D_%u7406%u5DE5_sex_%u5408%u80A5%u6821%u533A_%u7EDF%u62DB%u4E00%u6279'

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
    
    WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.CLASS_NAME, 'filter')))
    province_elements = driver.find_elements(By.XPATH, '/html/body/div/div[2]/div[2]/div[2]/div/div[2]/div/div[1]/dl/dd/a')
    province_len = len(province_elements)
    for province_id in range(province_len):
        province_elements = driver.find_elements(By.XPATH, '/html/body/div/div[2]/div[2]/div[2]/div/div[2]/div/div[1]/dl/dd/a')
        province_name = province_elements[province_id].text
        print('省份：', province_name)
        province_elements[province_id].click()
        
        WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, 'filterPlace')))
        year_elements = driver.find_elements(By.XPATH, '//*[@id="filterPlace"]/dl[1]/dd/a')
        year_len = len(year_elements)
        for year_id in range(year_len):
            year_elements = driver.find_elements(By.XPATH, '//*[@id="filterPlace"]/dl[1]/dd/a')
            year_name = year_elements[year_id].text
            print('年份：', year_name)
            year_elements[year_id].click()
            
            WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, 'filterPlace')))
            subject_elements = driver.find_elements(By.XPATH, '//*[@id="filterPlace"]/dl[2]/dd/a')
            subject_len = len(subject_elements)  
            for subject_id in range(subject_len):
                subject_elements = driver.find_elements(By.XPATH, '//*[@id="filterPlace"]/dl[2]/dd/a')
                subject_name = subject_elements[subject_id].text
                print('科类：', subject_name)
                subject_elements[subject_id].click()
                
                WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, 'filterPlace')))
                campus_elements = driver.find_elements(By.XPATH, '//*[@id="filterPlace"]/dl[3]/dd/a')
                campus_len = len(campus_elements)
                for campus_id in range(campus_len):
                    campus_elements = driver.find_elements(By.XPATH, '//*[@id="filterPlace"]/dl[3]/dd/a')
                    campus_name = campus_elements[campus_id].text
                    print('校区：', campus_name)
                    campus_elements[campus_id].click()
            
                    WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, 'filterPlace')))
                    type_elements = driver.find_elements(By.XPATH, '//*[@id="filterPlace"]/dl[4]/dd/a')
                    type_len = len(type_elements)
                    for type_id in range(type_len):
                        type_elements = driver.find_elements(By.XPATH, '//*[@id="filterPlace"]/dl[4]/dd/a')
                        type_name = type_elements[type_id].text
                        print('类型：', type_name)
                        type_elements[type_id].click()
    
                        # 爬取数据并存储
                        data = get_table_data(driver)
                        saveIntoCsv(data, province_name, year_name,driver)

def saveIntoCsv(data, province_name, year_name, driver):
    filename = f"{year_name}-{province_name}-合肥工业大学.csv"
    file_path = os.path.join(f'学校信息/合肥工业大学', filename)

    with open(file_path, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if os.path.getsize(file_path) == 0:  # 检查文件是否为空
            header_data = get_table_header(driver)
            writer.writerow(header_data)
        writer.writerows(data)



def get_table_header(driver):
    WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, 'sszygradeListPlace')))
    header = driver.find_element(By.ID, 'sszygradeListPlace')
    column_headers = header.find_elements(By.XPATH, '//*[@id="sszygradeListPlace"]/thead/tr/th')
    header_data = [column_header.text.strip() for column_header in column_headers]
    print("表格头部信息：", header_data)
    return header_data

def get_table_data(driver):
    WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, 'sszygradeListPlace')))
    data_wrapper = driver.find_element(By.ID, 'sszygradeListPlace')

    rows = data_wrapper.find_elements(By.XPATH, '//*[@id="sszygradeListPlace"]/tbody/tr')
    data = []
    for row in rows:
        try:
            cells = row.find_elements(By.XPATH, './td')
            row_data = []
            for cell in cells:
                row_data.append(cell.text.strip())
            data.append(row_data)
        except StaleElementReferenceException:
            rows = data_wrapper.find_elements(By.XPATH, '//*[@id="sszygradeListPlace"]/tbody/tr')
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
