__author__ = 'huangxingyue'

from selenium import webdriver
import numpy as np
import pandas as pd
import datetime
import time

data = []

driver = webdriver.Firefox()
driver.implicitly_wait(10)
driver.get('http://utc16.capitalrnts.com/student/orders.aspx?game_id=654')
elem = driver.find_element_by_id('ctl00_content_email')
elem.send_keys('xh890@nyu.edu')  # edit account name
elem = driver.find_element_by_id('ctl00_content_password')
elem.send_keys('ERA1108')  # edit password
elem = driver.find_element_by_id('ctl00_content_btnLogin')
elem.click()

driver.get('http://utc16.capitalrnts.com/student/orders.aspx?game_id=654')
elem = driver.find_element_by_link_text('Trading')
elem.click()

for tr in driver.find_elements_by_xpath('//tbody[@id="ctl00_content_tblBody"]//tr'):
    tds = tr.find_elements_by_tag_name('td')
    if tds:
        data.append([td.text for td in tds])

df = pd.DataFrame()
for row in data:
    arr = np.array(row).reshape((1, 11))
    df = df.append(pd.DataFrame(arr))
df = df.set_index(df.iloc[:, 0], drop=True)
df = df.iloc[:, [1]]
df.index.name = 'Securities'
df.columns = [datetime.datetime.now().time()]

driver.close()

def scrape():
    data = []

    driver = webdriver.Firefox()
    driver.implicitly_wait(10)
    driver.get('http://utc16.capitalrnts.com/student/orders.aspx?game_id=654')
    elem = driver.find_element_by_id('ctl00_content_email')
    elem.send_keys('xh890@nyu.edu')  # edit account name
    elem = driver.find_element_by_id('ctl00_content_password')
    elem.send_keys('ERA1108')  # edit password
    elem = driver.find_element_by_id('ctl00_content_btnLogin')
    elem.click()

    # elem = driver.find_element_by_xpath(".//a[@href='/student/games.aspx']")
    driver.get('http://utc16.capitalrnts.com/student/orders.aspx?game_id=654')
    elem = driver.find_element_by_link_text('Trading')
    elem.click()

    for tr in driver.find_elements_by_xpath('//tbody[@id="ctl00_content_tblBody"]//tr'):
        tds = tr.find_elements_by_tag_name('td')
        if tds:
            data.append([td.text for td in tds])
    temp = pd.DataFrame()
    for row in data:
        arr = np.array(row).reshape((1, 11))
        temp = temp.append(pd.DataFrame(arr))
    data2 = temp.iloc[:, [1]]
    df[datetime.datetime.now().time()] = data2.values
    driver.close()

for i in range(20):
    time.sleep(60)  # edit frequency
    scrape()

df.to_csv('data.csv')