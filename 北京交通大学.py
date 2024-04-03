from __future__ import annotations
import copy
import csv
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
import time


bind_address = "0.0.0.0"
wait_timeout = 100
poll_frequency = 0.2
tmp_dir = "/tmp"

logging_level = "DEBUG"
logging_format = "%(asctime)s --- %(levelname)s: %(message)s"
url = 'https://zsw.bjtu.edu.cn/list/index/id/37.html'

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
    
    WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, 'condition-province')))
    province_elements = driver.find_elements(By.XPATH, '//*[@id="condition-province"]/li/a')
    for province_element in province_elements:
        province_name = province_element.text
        print('省份：', province_name)
        province_element.click()
        time.sleep(2)

        WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, 'condition-year')))
        year_elements = driver.find_elements(By.XPATH, '//*[@id="condition-year"]/li/a')
        for year_element in year_elements:
            year_name = year_element.text
            print('年份：', year_name)
            year_element.click()
            time.sleep(2)
            
            WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, 'condition-school')))
            school_elements = driver.find_elements(By.XPATH, '//*[@id="condition-school"]/li/a')
            for school_element in school_elements:
                school_name = school_element.text
                print('校区：', school_name)
                school_element.click()
                time.sleep(2)

                WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, 'condition-sort')))
                kelei_elements = driver.find_elements(By.XPATH, '//*[@id="condition-sort"]/li/a')
                for kelei_element in kelei_elements:
                    kelei_name = kelei_element.text
                    print('科类：', kelei_name)
                    kelei_element.click()
                    time.sleep(2)

                    WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, 'condition-type')))
                    leixing_elements = driver.find_elements(By.XPATH, '//*[@id="condition-type"]/li/a')
                    for leixing_element in leixing_elements:
                        leixing_name = leixing_element.text
                        print('类型：', leixing_name)
                        leixing_element.click()
                        time.sleep(2)
                        data = get_table_data(driver,school_name,leixing_name)
                        saveIntoCsv(data, year_name,province_name,driver)

def saveIntoCsv(data, province_name, year_name, driver):
    filename = f"{year_name}-{province_name}-北京交通大学.csv"
    file_path = os.path.join(f'学校信息/北京交通大学', filename)

    with open(file_path, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if os.path.getsize(file_path) == 0:  # 检查文件是否为空
            header_data = get_table_header(driver)
            header_data.append("校区")
            header_data.append("类型")
            writer.writerow(header_data)
        writer.writerows(data)



def get_table_header(driver):
    WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, 'one-year-table')))
    header = driver.find_element(By.ID, 'one-year-table')
    column_headers = header.find_elements(By.XPATH, '//*[@id="one-year-table"]/thead/tr/th[position() != 5]')
    header_data = [column_header.text.strip() for column_header in column_headers]
    print("表格头部信息：", header_data)
    return header_data


def get_table_data(driver,school_name,leixing_name):
    WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, 'one-year-table')))
    data_wrapper = driver.find_element(By.ID, 'one-year-table')
    rows = data_wrapper.find_elements(By.XPATH, '//*[@id="one-year-table"]/tbody/tr')
    data = []
    for row in rows:
        try:
            cells = row.find_elements(By.XPATH, './td[position() != 5]')
            row_data = []
            for cell in cells:
                row_data.append(cell.text.strip())
            row_data.append(school_name)
            row_data.append(leixing_name)    
            data.append(row_data)
        except StaleElementReferenceException:
            rows = data_wrapper.find_elements(By.XPATH, '//*[@id="one-year-table"]/tbody/tr')
            continue
    print("表格数据：", data)
    return data

def run(logger: logging.Logger):
    global _logger
    _logger = logger
    try:
        _driver = restart_session('headless')
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
