from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
import time
import os
import shutil
import csv

chrome_options = webdriver.ChromeOptions()
'''chrome_options.add_argument('--headless')
chrome_options.add_argument('--disable-gpu')'''

driver = webdriver.Chrome(service=ChromeService(), options=chrome_options)
driver.get('https://codis.cwa.gov.tw/StationData')
driver.set_window_size(1660, 945)
driver.implicitly_wait(5)

time.sleep(2)
#改成清單模式，並選擇只看北部地區
driver.find_element(By.CSS_SELECTOR, '#switch_display > button:nth-child(2)').click()
driver.find_element(By.CSS_SELECTOR, '#station_filter > div > section > ul > li:nth-child(3) > div > div.col-7.px-0 > select > option:nth-child(2)').click()

#選擇站點
times = driver.find_element(By.CSS_SELECTOR, '#station_count').text
for nowtimes in range(1, int(times)+1): #用 range 開始的數字來調整從哪個測站開始
    name = driver.find_element(By.CSS_SELECTOR, '#station_table > table > tbody > tr:nth-child({}) > td:nth-child(1) > div'.format(nowtimes)).text
    if('雷達站' in name):
        print("雷達站跳過")
        continue
    print("目前測站{}，還有{}個測站".format(name, str(int(times)-nowtimes)))
    number = driver.find_element(By.CSS_SELECTOR, '#station_table > table > tbody > tr:nth-child({}) > td:nth-child(2) > div'.format(nowtimes)).text
    driver.find_element(By.CSS_SELECTOR, '#station_table > table > tbody > tr:nth-child({}) > td:nth-child(10) > div'.format(nowtimes)).click()

    file_path = os.path.join(os.getcwd(), '{}({})'.format(name, number))
    file_list = os.listdir(file_path)
    flag = False
    for file_name in file_list:
        if("xlsx" in file_name):
            continue
        csvfile = open(os.path.join(file_path, file_name), encoding='utf-8') # 開啟 CSV 檔案
        raw_data = csv.reader(csvfile) # 讀取 CSV 檔案
        data = list(raw_data) # 轉換成二維串列
        csvfile.close()
        if(len(data) > 2):
            continue
        date = list(map(int, file_name.split('.')[0].split('-')[1:]))
        print("測站：{}，年月日：{}".format(name, date))
        flag = True

        #修改觀測年月
        driver.find_element(By.CSS_SELECTOR, '#main_content > section.lightbox-tool > div > div > section > div:nth-child(5) > div.lightbox-tool-type-ctrl > div.lightbox-tool-type-ctrl-form > label > div > div.vdatetime > input').click()
        time.sleep(1)

        #年
        nyear = driver.find_element(By.CSS_SELECTOR, '#main_content > section.lightbox-tool > div > div > section > div:nth-child(5) > div.lightbox-tool-type-ctrl > div.lightbox-tool-type-ctrl-form > label > div > div.vdatetime > div > div.vdatetime-popup > div.vdatetime-popup__header > div.vdatetime-popup__year').text
        driver.find_element(By.CSS_SELECTOR, '#main_content > section.lightbox-tool > div > div > section > div:nth-child(5) > div.lightbox-tool-type-ctrl > div.lightbox-tool-type-ctrl-form > label > div > div.vdatetime > div > div.vdatetime-popup > div.vdatetime-popup__header > div.vdatetime-popup__year').click()
        driver.find_element(By.CSS_SELECTOR, '#main_content > section.lightbox-tool > div > div > section > div:nth-child(5) > div.lightbox-tool-type-ctrl > div.lightbox-tool-type-ctrl-form > label > div > div.vdatetime > div > div.vdatetime-popup > div.vdatetime-popup__body > div > div > div:nth-child({})'.format(100-(int(nyear)-date[0]-1))).click()

        now = driver.find_element(By.CSS_SELECTOR, '#main_content > section.lightbox-tool > div > div > section > div:nth-child(5) > div.lightbox-tool-type-ctrl > div.lightbox-tool-type-ctrl-form > label > div > div.vdatetime > div > div.vdatetime-popup > div.vdatetime-popup__header > div.vdatetime-popup__date').text
        #月
        nmin = now.split('月')[0]
        driver.find_element(By.CSS_SELECTOR, '#main_content > section.lightbox-tool > div > div > section > div:nth-child(5) > div.lightbox-tool-type-ctrl > div.lightbox-tool-type-ctrl-form > label > div > div.vdatetime > div > div.vdatetime-popup > div.vdatetime-popup__header > div.vdatetime-popup__date').click()
        if(date[1] == nmin):
            driver.find_element(By.CSS_SELECTOR, '#main_content > section.lightbox-tool > div > div > section > div:nth-child(5) > div.lightbox-tool-type-ctrl > div.lightbox-tool-type-ctrl-form > label > div > div.vdatetime > div > div.vdatetime-popup > div.vdatetime-popup__body > div > div > div.vdatetime-month-picker__item.vdatetime-month-picker__item--selected').click()
        else:
            driver.find_element(By.CSS_SELECTOR, '#main_content > section.lightbox-tool > div > div > section > div:nth-child(5) > div.lightbox-tool-type-ctrl > div.lightbox-tool-type-ctrl-form > label > div > div.vdatetime > div > div.vdatetime-popup > div.vdatetime-popup__body > div > div > div:nth-child({})'.format(date[1])).click()

        #日
        nday = now.split('月')[1].replace('日', '')
        if(date[2] == nday):
            driver.find_element(By.CSS_SELECTOR, '#main_content > section.lightbox-tool > div > div > section > div:nth-child(5) > div.lightbox-tool-type-ctrl > div.lightbox-tool-type-ctrl-form > label > div > div.vdatetime > div > div.vdatetime-popup > div.vdatetime-popup__body > div > div.vdatetime-calendar__month > div.vdatetime-calendar__month__day.vdatetime-calendar__month__day--selected').click()
        else:
            for i in range(8, 43):
                tmp = driver.find_element(By.CSS_SELECTOR, '#main_content > section.lightbox-tool > div > div > section > div:nth-child(5) > div.lightbox-tool-type-ctrl > div.lightbox-tool-type-ctrl-form > label > div > div.vdatetime > div > div.vdatetime-popup > div.vdatetime-popup__body > div > div.vdatetime-calendar__month > div:nth-child({})'.format(i))
                if(tmp.text == str(date[2])):
                    tmp.click()
                    break

        driver.find_element(By.CSS_SELECTOR, '#main_content > section.lightbox-tool > div > div > section > div:nth-child(5) > div.lightbox-tool-type-ctrl > div.lightbox-tool-type-ctrl-form > label > div > div.vdatetime > div > div.vdatetime-popup > div.vdatetime-popup__actions > div.vdatetime-popup__actions__button.vdatetime-popup__actions__button--confirm').click()

        time.sleep(1)
        driver.find_element(By.CSS_SELECTOR, '#main_content > section.lightbox-tool > div > div > section > div:nth-child(5) > div.lightbox-tool-type-ctrl > div.lightbox-tool-type-ctrl-btn-group > div').click()

    if(flag):
        time.sleep(2)
    filelist = os.listdir('C:\\Users\\GIGABYTE\\Downloads')
    for filename in filelist:
        if(number in filename):
            shutil.move('C:\\Users\\GIGABYTE\\Downloads\\' + filename, "C:\\Users\\GIGABYTE\\桌面\\VScode_Code\\天氣預測資料\\署屬有人站\\" + '{}({})\\'.format(name, number) + filename)
    
    driver.find_element(By.CSS_SELECTOR, '#main_content > section.lightbox-tool > div > header > div.lightbox-tool-close').click()
driver.quit()