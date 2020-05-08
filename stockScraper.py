import selenium
import time
from selenium import webdriver
from collections import defaultdict
#from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd


chrome_options = Options()
# maximized window
chrome_options.add_argument("--start-maximized")

# Use Chrome to access web
# k = int(input('Max Limit (usually not more than 2000): '))
# br = input('Branch (CS, EC, EE, ME, CV, BT): ')
chrome = webdriver.Chrome('/bin/chromedriver/chromedriver', chrome_options=chrome_options)


# Create an empty (for now) database to store students
companies = dict()

nameList = ['Reliance', 'TCS', 'ONGC', 'HDFC', 'Kotak Mahindra', 'HDFC Bank', 'Infosys', 'HCL Tech', 'Tata Steel', 'JSW Steel', 'UltraTechCement', 'Ambuja Cements', 'Bajaj Auto', 'Maruti Suzuki', 'M&M', 'HUL', 'GAIL'
, 'Power Finance', 'Cipla', 'Glenmark', 'BEL', 'Biocon', 'Syngen', 'Bharti Airtel', 'ITC', 'L&T Infotech', 'Larsen', 'Siemens', 'Havells India', 'Bajaj Finance']
  
chrome.get('https://www.moneycontrol.com/stocks/marketinfo/marketcap/bse/index.html')
#time.sleep(3)
#chrome.maximize_window()
k = 1
for num in range (2,101):
    name = chrome.find_element_by_xpath(r'//*[@id="mc_mainWrapper"]/div[3]/div[1]/div[10]/div[2]/div/table/tbody/tr[{}]/td[1]/a/b'.format(str(num))).text
    if name not in nameList:
        print(name, 'not required.')
        continue
    companies[name] = dict()
    time.sleep(0.2)
    companies[name]['Net Profit'] = ''
    companies[name]['Market Cap'] = chrome.find_element_by_xpath(r'//*[@id="mc_mainWrapper"]/div[3]/div[1]/div[10]/div[2]/div/table/tbody/tr[{}]/td[6]'.format(str(num))).text
    companies[name]['Last Price'] = chrome.find_element_by_xpath(r'//*[@id="mc_mainWrapper"]/div[3]/div[1]/div[10]/div[2]/div/table/tbody/tr[{}]/td[2]'.format(str(num))).text
    companies[name]['52 Week High'] = chrome.find_element_by_xpath(r'//*[@id="mc_mainWrapper"]/div[3]/div[1]/div[10]/div[2]/div/table/tbody/tr[{}]/td[4]'.format(str(num))).text
    companies[name]['52 Week Low'] = chrome.find_element_by_xpath(r'//*[@id="mc_mainWrapper"]/div[3]/div[1]/div[10]/div[2]/div/table/tbody/tr[{}]/td[5]'.format(str(num))).text
    chrome.find_element_by_xpath(r'//*[@id="mc_mainWrapper"]/div[3]/div[1]/div[10]/div[2]/div/table/tbody/tr[{}]/td[1]/a'.format(str(num))).click()
    print(str(k)+'.', 'Clicked on', name)
    try:
        chrome.find_element_by_xpath(r'//*[@id="mc_mainWrapper"]/div[3]/div[1]/div[10]/div[2]/div/table/tbody/tr[{}]/td[1]/a'.format(str(num))).click()
    except:
        pass
    WebDriverWait(chrome, 30).until(EC.presence_of_element_located((By.ID, "sec_valul")))
    time.sleep(0.5)
    companies[name]['Face Value'] = chrome.find_element_by_xpath(r'//*[@id="standalone_valuation"]/ul/li[3]/ul/li[3]/div[2]').text
    time.sleep(0.5)
    companies[name]['Book Value'] = chrome.find_element_by_xpath(r'//*[@id="standalone_valuation"]/ul/li[1]/ul/li[3]/div[2]').text
    time.sleep(0.5)
    companies[name]['P/E Ratio'] = chrome.find_element_by_xpath(r'//*[@id="standalone_valuation"]/ul/li[1]/ul/li[2]/div[2]').text
    time.sleep(0.5)
    companies[name]['Dividend %'] = chrome.find_element_by_xpath(r'//*[@id="standalone_valuation"]/ul/li[1]/ul/li[4]/div[2]').text
    k+=1
    time.sleep(1)
    try:
        companies[name]['Sector'] = chrome.find_element_by_xpath(r'//*[@id="sec_quotes"]/div[2]/div/div[1]/span[6]/a').text
    except:
        companies[name]['Sector'] = ''
    #time.sleep(3)
    #//*[@id="sec_quotes"]/div[2]/div/div[1]/span[5]
    time.sleep(0.5)
    chrome.execute_script("window.history.go(-1)")
    print(companies[name])
    WebDriverWait(chrome, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="mc_mainWrapper"]/div[3]/div[1]/div[10]/div[2]/div/table/tbody/tr[2]/td[6]')))

chrome.get('https://www.moneycontrol.com/stocks/marketinfo/netprofit/bse/index.html')
for num in range (2,101):
    name = chrome.find_element_by_xpath(r'//*[@id="mc_mainWrapper"]/div[3]/div[1]/div[10]/div[2]/div/table/tbody/tr[{}]/td[1]/a/b'.format(str(num))).text
    try:
        companies[name]['Net Profit'] = chrome.find_element_by_xpath(r'//*[@id="mc_mainWrapper"]/div[3]/div[1]/div[10]/div[2]/div/table/tbody/tr[{}]/td[5]'.format(str(num))).text
    except:
        pass
        

#print(companies)

writer = pd.ExcelWriter('Company Stats.xlsx', engine='xlsxwriter')
pd.DataFrame(companies).T.to_excel(writer, sheet_name='Top 100')
workbook  = writer.book
worksheet = writer.sheets['Top 100']

try:
    format1 = workbook.add_format({'align': 'centre'})
    worksheet.set_column('A:A', 20, format1)
    worksheet.set_column('B:B', 14, format1)
    worksheet.set_column('C:C', 16, format1)
    worksheet.set_column('D:D', 14, format1)
    worksheet.set_column('E:F', 18, format1)
    worksheet.set_column('G:I', 16, format1)
    worksheet.set_column('J:J', 18, format1)
    worksheet.set_column('K:K', 36, format1)

except:
    format1 = workbook.add_format({'align': 'center'})
    worksheet.set_column('A:A', 20, format1)
    worksheet.set_column('B:B', 14, format1)
    worksheet.set_column('C:C', 16, format1)
    worksheet.set_column('D:D', 14, format1)
    worksheet.set_column('E:F', 18, format1)
    worksheet.set_column('G:I', 16, format1)
    worksheet.set_column('J:J', 18, format1)
    worksheet.set_column('K:K', 36, format1)

writer.save()



