from bs4 import BeautifulSoup
from io import BytesIO
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import Select
from time import sleep

import csv
import datetime
import numpy as np
import pandas as pd
import requests
import time

def waiting(b): 
    msg = b.find_element_by_id("divBlock")
    while msg.is_displayed():
        sleep(1)

#STEP 1: 呼叫Chrome的Web driver
main_driver = webdriver.Chrome('/Users/pptaira/Documents/Fintech_2019S/chromedriver') 
main_driver.implicitly_wait(20)

#STEP 2: 打開目標網頁
main_driver.get('https://www.sitca.org.tw/ROC/Industry/IN2201.aspx?pid=IN2221_01&fbclid=IwAR02IyRfpqDKlVtnzOAtMwPriqP7BJ0_zLtBpexmMdgKx5P1SbMxnVcqokQ')
waiting(main_driver)
select = Select(main_driver.find_element_by_id('ctl00_ContentPlaceHolder1_ddlQ_COLUMN'))
select.select_by_visible_text("定時定額")

select = Select(main_driver.find_element_by_id('ctl00_ContentPlaceHolder1_ddlQ_YM'))
options = len(select.options)
for i in range(216, options):
    select_rec = Select(main_driver.find_element_by_id('ctl00_ContentPlaceHolder1_ddlQ_YM'))
    waiting(main_driver)
    select_rec.select_by_index(i)
    waiting(main_driver)
    ele = main_driver.find_element_by_xpath('//*[@id="ctl00_ContentPlaceHolder1_BtnQuery"]')
    ele.click()
    waiting(main_driver)

    #STEP 3: 抓取表格資料
    soup = BeautifulSoup(main_driver.page_source, "html.parser")
    tables = soup.findAll('table')
    df = pd.read_html(str(tables))[3]
    
    df.to_csv(str(i) + '.csv', encoding = 'utf-8')
    
main_driver.close()