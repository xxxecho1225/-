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
from selenium.webdriver.support.ui import Select

bind_address = "0.0.0.0"
wait_timeout = 10
poll_frequency = 0.2
tmp_dir = "/tmp"

logging_level = "DEBUG"
logging_format = "%(asctime)s --- %(levelname)s: %(message)s"
url = 'https://lqcx.njfu.edu.cn/page/enrollment-plan.html?flag=2'

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
        
    time.sleep(3)
    
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'yearInfo')))  
    # 找到年份下拉框元素
    year_select = Select(driver.find_element(By.ID, 'yearInfo'))
    # 获取所有年份选项
    year_options = year_select.options
    # 提取每个年份的值
    years = [option.get_attribute('value') for option in year_options]
    year_len = len(years)
    for year_id in range(year_len):
        year_selects = Select(driver.find_element(By.ID, 'yearInfo'))
        year_name = year_selects[year_id].text
        print('年份：', year_name)
        time.sleep(3)
        driver.execute_script("arguments[0].selectedIndex = {}; arguments[0].dispatchEvent(new Event('change'))".format(year_id), year_selects[year_id])
    
        time.sleep(3)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'province')))  
    # 找到年份下拉框元素
        province_select = Select(driver.find_element(By.ID, 'province'))
        # 获取所有年份选项
        province_options = province_select.options
        # 提取每个年份的值
        years = [option.get_attribute('value') for option in year_options]
        province_len = len(province_options)
        for province_id in range(province_len):
            province_selects = Select(driver.find_element(By.ID, 'province'))
            province_name = province_select[province_id].text
            print('省份：', province_name)
            driver.execute_script("arguments[0].selectedIndex = {}; arguments[0].dispatchEvent(new Event('change'))".format(province_id), province_selects[province_id])
            clickbutton(driver)
            data = get_table_data(driver)
            saveIntoXlsx(data, province_name, year_name,driver)
def clickbutton(driver):
    button = WebDriverWait(driver, 200).until(EC.presence_of_element_located((By.CLASS_NAME, 'btn')))
    button.click()
def saveIntoXlsx(data, province_name, year_name,driver):

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
    filename = f"{year_name}-{province_name}-南京林业大学.csv"
    os.makedirs('学校信息\南京林业大学', exist_ok=True)
    wb.save(os.path.join('学校信息\南京林业大学', filename))

def get_table_header(driver):
    # 等待表格数据包装器元素出现
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'result-title')))  
    
    # 找到表头元素
    table = driver.find_element(By.CLASS_NAME, 'result-title')

    headers = table.find_elements(By.XPATH, '//div[@class="result-title"]/ul/li')
    
    # 提取列头文本
    header_data = [column_header.text.strip() for column_header in headers]
    
    print("表格头部信息：", header_data)
    
    return header_data


def get_table_data(driver):
    WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, 'contentInfoId')))
    table = driver.find_element(By.ID, 'contentInfoId')
    # 使用XPath选择所有td元素
    data = []
    i = 0
    while True:
        try:
            row = table.find_elements(By.XPATH, '//*[@id="contentInfoId"]/ul/li')[i]
            row_data = [cell.get_attribute('innerText') for cell in row]
            if row_data:
                data.append(row_data)
            i += 1
        except IndexError:
            break
        except StaleElementReferenceException:
            # 页面元素已过时，尝试重新查找表格元素
            table = driver.find_element(By.ID, 'contentInfoId')
            continue
    print("2222222",data)
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
