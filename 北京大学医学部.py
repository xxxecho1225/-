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
url = 'https://bkzs.bjmu.edu.cn/zsxx/lnfs/index.htm'

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
    
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/article/div/section/div/div/div/ul/li[7]/a'))) 
    # 定位省份元素并循环
    province_elements = driver.find_elements(By.XPATH, '/html/body/article/div/section/div/div/div/ul/li[7]/a')
    for province_element in province_elements:
        while True:
            try:
                province_name = province_element.text
                print('名字：', province_name)
                province_element.click()
                break
            except StaleElementReferenceException:
                # 如果发生异常，继续尝试
                province_elements = driver.find_elements(By.XPATH, '/html/body/article/div/section/div/div/div/ul/li[7]/a')
                continue
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/article/div[2]/section/div/div[2]/div[2]/table[1]/tbody/tr[1]/td/span/span/strong'))) 
        # 定位省份元素并循环
        table_elements = driver.find_elements(By.XPATH, '/html/body/article/div[2]/section/div/div[2]/div[2]/table[1]/tbody/tr[1]/td/span/span/strong')
        for table_element in table_elements:
            table_name = table_element.text
            print('表格1标题名字:', table_name)
            table_element.click()
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/article/div[2]/section/div/div[2]/div[2]/table[2]/tbody/tr[1]/td/span/span/strong'))) 
            # 定位省份元素并循环
            table2_elements = driver.find_elements(By.XPATH, '/html/body/article/div[2]/section/div/div[2]/div[2]/table[2]/tbody/tr[1]/td/span/span/strong')
            for table2_element in table2_elements:
                table2_name = table2_element.text
                print('表格2标题名字:', table2_name)
                table2_element.click()

                data = get_table_data(driver)
                # 获取第二个表格数据
                data2 = get_table_data2(driver)

                # 创建工作表并写入数据
                wb = Workbook()
                ws1 = wb.active
                ws1.title = table_name
                ws2 = wb.create_sheet(title=table2_name)
                for row in data:
                    ws1.append(row)
                for row in data2:
                    ws2.append(row)

                # 保存工作簿到文件
                filename = f"{province_name}.xlsx"
                os.makedirs('学校信息\北京大学医学部', exist_ok=True)
                wb.save(os.path.join('学校信息\北京大学医学部', filename))



def get_table_data(driver):
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/article/div[2]/section/div/div[2]/div[2]/table[1]/tbody/tr')))
    
    # 找到表格数据包装器元素
    data_wrappers = driver.find_elements(By.XPATH, '/html/body/article/div[2]/section/div/div[2]/div[2]/table[1]/tbody/tr')

    data = []

 # 提取每行数据
    for row in data_wrappers:
        try:
            # 找到当前行下的所有td元素
            cells = row.find_elements(By.XPATH, './td')
            row_data = []
            # 提取每个td中的div的文本内容
            for cell in cells:
                row_data.append(cell.text.strip())
            data.append(row_data)
        except StaleElementReferenceException:
            # 如果发生异常，则重新定位元素
            data_wrappers = driver.find_elements(By.XPATH, '/html/body/article/div[2]/section/div/div[2]/div[2]/table[1]/tbody/tr')
            continue

    print("表格数据：", data)

    
    return data

def get_table_data2(driver):
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '/html/body/article/div[2]/section/div/div[2]/div[2]/table[2]/tbody/tr')))
    
    # 找到表格数据包装器元素
    data_wrappers = driver.find_elements(By.XPATH, '/html/body/article/div[2]/section/div/div[2]/div[2]/table[2]/tbody/tr')
    
    
    # 找到所有行数据
    data2 = []

 # 提取每行数据
    for row in data_wrappers:
        try:
            # 找到当前行下的所有td元素
            cells = row.find_elements(By.XPATH, './td')
            row_data = []
            # 提取每个td中的div的文本内容
            for cell in cells:
                row_data.append(cell.text.strip())
            data2.append(row_data)
        except StaleElementReferenceException:
            # 如果发生异常，则重新定位元素
            data_wrappers = driver.find_elements(By.XPATH, '/html/body/article/div[2]/section/div/div[2]/div[2]/table[2]/tbody/tr')
            continue

    print("表格数据2:", data2)

    
    return data2

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
