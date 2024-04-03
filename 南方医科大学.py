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
url = 'https://portal.smu.edu.cn/bkzs/bkzn/wnfs.htm'

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
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="scArea1"]/a')))
    
    # 定位省份元素并循环
    province_elements = driver.find_elements(By.XPATH, '//*[@id="scArea1"]/a')
    for province_element in province_elements:
        province_name = province_element.text
        print('省份：', province_name)
        province_element.click()
        
        # 定位年份元素并循环
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="scArea2"]/a')))
        year_elements = driver.find_elements(By.XPATH, '//*[@id="scArea2"]/a')
        for year_element in year_elements:
            year_name = year_element.text
            print("年份名字：",year_name)
            year_element.click()
            
            data_all = {}  # 创建一个新的字典来存储数据
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="scArea3"]/a')))
            subject_elements = driver.find_elements(By.XPATH, '//*[@id="scArea3"]/a')
            for subject_element in subject_elements:
                subject_name = subject_element.text
                print('科类：', subject_name)
                subject_element.click()
                
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="scArea4"]/a')))
                type_elements = driver.find_elements(By.XPATH, '//*[@id="scArea4"]/a')
                for type_element in type_elements:
                    type_name = type_element.text
                    print('类型：', type_name)
                    type_element.click()
        
                    data = get_table_data(driver)
                    data_all[f'{subject_name}_{type_name}'] = copy.deepcopy(data)
                    
                # 保存当前年份和省份的数据到文件
                saveIntoCsv(data_all, province_name, year_name, driver)


def saveIntoCsv(data_all,province_name, year_name,driver):
    wb = Workbook()
    
    # 删除默认创建的工作表
    ws = wb.active
    wb.remove(ws)
    
    # 保存数据到各个工作表
    for sheet_name, data in data_all.items():
        # 创建一个新的工作表
        ws = wb.create_sheet(title=sheet_name)
        
        # 写入表头数据
        header_data = get_table_header(driver)
        ws.append(header_data)
        
        # 写入表格数据
        for row in data:
            ws.append(row)
    # 保存工作簿到文件
    filename = f"{year_name}-{province_name}-南方医科大学.csv"
    os.makedirs('学校信息\南方医科大学', exist_ok=True)
    wb.save(os.path.join('学校信息\南方医科大学', filename))

def get_table_header(driver):
    WebDriverWait(driver, 100).until(EC.presence_of_element_located((By.CLASS_NAME, 'el-table__header-wrapper')))
    
    # 找到表头元素
    header = driver.find_element(By.CLASS_NAME, 'el-table__header-wrapper')
    
    WebDriverWait(header, 100).until(EC.presence_of_element_located((By.CLASS_NAME, 'el-table__header')))
    
    # 找到表头元素
    header2 = header.find_element(By.CLASS_NAME, 'el-table__header')
    # 找到所有的列头元素
    column_headers = header2.find_elements(By.XPATH, '//*[@id="searchForm1074188"]/div/div/div/div/div/div/table/thead/tr/th')

    
    # 提取列头文本
    header_data = [column_header.text.strip() for column_header in column_headers]
    
    print("表格头部信息：", header_data)
    
    return header_data

def get_table_data(driver):
    # 等待表格数据加载完成
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'el-table__body')))
    
    # 找到表格数据元素
    table_body = driver.find_element(By.CLASS_NAME, 'el-table__body')
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="searchForm1074188"]/div/div/div/div/div/div/table/tbody/tr')))
    # 找到所有行数据
    rows = table_body.find_elements(By.XPATH, '//*[@id="searchForm1074188"]/div/div/div/div/div/div/table/tbody/tr')
    data = []

# 提取每行数据
    for row in rows:
        try:
            # 找到当前行下的所有td元素
            cells = row.find_elements(By.XPATH, './td')
            row_data = []
            # 提取每个td中的div的文本内容
            for cell in cells:
                div = cell.find_element(By.XPATH, './div')
                row_data.append(div.text.strip())
            data.append(row_data)
        except StaleElementReferenceException:
            # 如果发生异常，则重新定位元素
            rows = table_body.find_elements(By.XPATH, '//*[@id="searchForm1074188"]/div/div/div/div/div/div/table/tbody/tr')
            continue
    print("表格信息：", data)

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
