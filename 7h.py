__author__ = 'huangxingyue'

from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
import numpy as np
import pandas as pd
import time
from xlwt.Workbook import *

# store data in dictionary
overview = {}
ori_data = {}

RANK = 4  # set the number of top traders to be researched

driver = webdriver.Firefox()
driver.implicitly_wait(15)
driver.get("http://trader.7hcn.com")

WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, '//*[@title="股指"]')))
elem = driver.find_element_by_xpath('//*[@title="股指"]')
elem.click()

WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, '//tbody[@id="mytable"]//tr')))
time.sleep(5)
cells = driver.find_elements_by_xpath('//tbody[@id="mytable"]//tr')[1:RANK+1]
keys = [item.text.split('\n')[1] for item in cells]

for i in range(RANK):
    elems = driver.find_elements_by_xpath('//tbody[@id="mytable"]//td[@align="left"]')[1:RANK*2+1]
    key = keys[i]
    elems[2*i].click()
    # DATA1-OVERVIEW
    WebDriverWait(driver, 180).until(EC.element_to_be_clickable((By.XPATH, '//*[@todo="overview"]')))
    elem3 = driver.find_element_by_xpath('//*[@todo="overview"]')
    elem3.click()

    WebDriverWait(driver, 180).until(EC.element_to_be_clickable((By.XPATH, '//table[@class="tableborder"][@align="center"]//tbody')))
    data1 = []
    for tr in driver.find_elements_by_xpath('//table[@class="tableborder"][@align="center"]//tbody'):
        tds = tr.find_elements_by_tag_name('td')
        if tds:
            data1.append([td.text for td in tds])
    data1 = data1[0][:-1]
    data1 = np.array(data1).reshape(int(len(data1)/2), 2)
    df1 = pd.DataFrame(data1)
    overview[keys[i]] = df1

    # DATA2-ORIGINAL TRANSACTION DATA
    WebDriverWait(driver, 180).until(EC.element_to_be_clickable((By.XPATH, '//a[@todo="org_data"]')))
    elem4 = driver.find_element_by_xpath('//a[@todo="org_data"]')
    elem4.click()
    WebDriverWait(driver, 180).until(EC.element_to_be_clickable((By.XPATH, '//td[@colspan="15"]//div[@class="pages"]//a[@class="last"]')))
    page_num = int(driver.find_element_by_xpath('//td[@colspan="15"]//div[@class="pages"]//a[@class="last"]').text.split(' ')[-1])

    data2 = []
    ths = driver.find_elements_by_xpath('//div[@id="dialog_result"]//table[@class="tableborder"]//th')
    data2.append([th.text for th in ths])

    for j in range(page_num-2):
        for tr in driver.find_elements_by_xpath('//table[@class="tableborder"]//tbody[@id="source_list"]'):
            tds = tr.find_elements_by_tag_name('td')
            if tds:
                data2.append([td.text for td in tds][:-1])
        elem = driver.find_element_by_xpath('//td[@colspan="15"]//div[@class="pages"]//a[@class="next"]')
        elem.click()
        WebDriverWait(driver, 180).until(EC.staleness_of(elem))
        WebDriverWait(driver, 180).until(EC.element_to_be_clickable((By.XPATH, '//td[@colspan="15"]//div[@class="pages"]//a[@class="next"]')))

    for tr in driver.find_elements_by_xpath('//table[@class="tableborder"]//tbody[@id="source_list"]'):
        tds = tr.find_elements_by_tag_name('td')
        if tds:
            data2.append([td.text for td in tds][:-1])

    df2 = pd.DataFrame()
    for row in data2:
        row_num = int(len(row)/15)
        narray = np.array(row).reshape((row_num, 15))
        df2 = df2.append(pd.DataFrame(narray))

    ori_data[keys[i]] = df2

    # output results in excel with two sheets
    wb = Workbook()
    ws1 = wb.add_sheet('overview')
    ws2 = wb.add_sheet('ori_data')
    writer = pd.ExcelWriter('{}.xlsx'.format(keys[i]))
    df1.to_excel(writer, 'overview', header=False, index=False)
    df2.to_excel(writer, 'ori_data', header=False, index=False)
    writer.save()

    elem = driver.find_element_by_xpath('//a[@title="close"]')
    elem.click()
    driver.close()

    driver = webdriver.Firefox()
    driver.implicitly_wait(15)
    driver.get("http://trader.7hcn.com")
    WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, '//*[@title="股指"]')))
    elem = driver.find_element_by_xpath('//*[@title="股指"]')
    elem.click()
    time.sleep(10)

driver.close()


# DATA3-GRAPHIC DATA
