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


bind_address = "0.0.0.0"
wait_timeout = 10
poll_frequency = 0.2
tmp_dir = "/tmp"

logging_level = "DEBUG"
logging_format = "%(asctime)s --- %(levelname)s: %(message)s"
url = 'https://yx.lzu.edu.cn/lzuzsb/stuweb/scoreweb/score.jsp'

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
    
    WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, 'sfselect')))
    sf_elements = driver.find_elements(By.XPATH, '//*[@id="sfselect"]/option[text()="安徽" or text() = "福建" or text() = "北京" or text() = "甘肃"or text() = "广东"or text() = "广西" or text() = "贵州"or text() = "海南"or text() = "河北" or text() = "河南"or text() = "黑龙江"or text() = "湖北"or text() = "湖南"or text() = "吉林"or text() = "江苏"or text() = "江西"or text() = "辽宁"or text() = "内蒙古" or text() = "宁夏"or text() = "青海"or text() = "山东"or text() = "山西"or text() = "陕西" or text() = "上海"or text() = "四川"or text() = "天津"or text() = "新疆"or text() = "云南"or text() = "浙江" or text() = "重庆"]')
    sf_len = len(sf_elements)
    for sf_id in range(sf_len):
        sf_elements = driver.find_elements(By.XPATH, '//*[@id="sfselect"]/option[text()="安徽" or text() = "福建" or text() = "北京" or text() = "甘肃"or text() = "广东"or text() = "广西" or text() = "贵州"or text() = "海南"or text() = "河北" or text() = "河南"or text() = "黑龙江"or text() = "湖北"or text() = "湖南"or text() = "吉林"or text() = "江苏"or text() = "江西"or text() = "辽宁"or text() = "内蒙古" or text() = "宁夏"or text() = "青海"or text() = "山东"or text() = "山西"or text() = "陕西" or text() = "上海"or text() = "四川"or text() = "天津"or text() = "新疆"or text() = "云南"or text() = "浙江" or text() = "重庆"]')
        sf_name = sf_elements[sf_id].text
        print('省份：', sf_name)
        sf_elements[sf_id].click()
    
        WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.ID, 'nfselect')))
        year_elements = driver.find_elements(By.XPATH, '//*[@id="nfselect"]/option[text()="2023年" or text() = "2022年" or text() = "2021年"]')
        year_len = len(year_elements)
        for year_id in range(year_len):
            year_elements = driver.find_elements(By.XPATH, '//*[@id="nfselect"]/option[text()="2023年" or text() = "2022年" or text() = "2021年"]')
            year_name = year_elements[year_id].text
            print('年份：', year_name)
            year_elements[year_id].click()
            
            clickbutton(driver)

            time.sleep(3)
            data = get_table_data(driver)
            save_into_csv(data, year_name,sf_name,driver)

def save_into_csv(data, year_name,sf_name, driver):
    wb = Workbook()
    # 首先删除默认创建的工作表
    default_ws = wb.active
    wb.remove(default_ws)

    # 创建工作表并写入数据
    ws = wb.create_sheet(title="Data")
    
    # 添加表头
    header_data = get_table_header(driver)
    ws.append(header_data)
    
    # 写入数据
    for row in data:
        ws.append(row)
    
    # 保存文件
    filename = f"{year_name}-{sf_name}-兰州大学.csv"
    os.makedirs('学校信息\兰州大学', exist_ok=True)
    wb.save(os.path.join('学校信息\兰州大学', filename))


    
def clickbutton(driver):
    button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="searchButton"]')))
    # 点击按钮
    button.click()

def get_table_header(driver):
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'tab1')))
    header = driver.find_element(By.ID, 'tab1')
    column_headers = header.find_elements(By.XPATH, '//*[@id="tab1"]/thead/tr[1]/th')
    header_data = [column_header.text.strip() for column_header in column_headers]
    print("表格头部信息：", header_data)
    return header_data


def get_table_data(driver):
    WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.ID, 'tab1')))
    data_wrapper = driver.find_element(By.ID, 'tab1')
    rows = data_wrapper.find_elements(By.XPATH, '//*[@id="tab1"]/thead/tr')
    data = []
    for row in rows:
        try:
            cells = row.find_elements(By.XPATH, './td')
            row_data = []
            for cell in cells:
                row_data.append(cell.text.strip())
            if row_data:    
                data.append(row_data)
        except StaleElementReferenceException:
            rows = data_wrapper.find_elements(By.XPATH, '//*[@id="tab1"]/thead/tr')
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
