##### Imports #####
from selenium import webdriver
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import pandas as pd
from datetime import date
from openpyxl import load_workbook

##### Comments #####

#path is where the chromedriver is on my local computer - so that needs to be changed for everyone
#https://chromedriver.chromium.org/getting-started
#if on mac and can't chromedriver: https://stackoverflow.com/questions/60362018/macos-catalinav-10-15-3-error-chromedriver-cannot-be-opened-because-the-de

#driver = webdriver.Chrome("/Users/serena/Downloads/chromedriver") 
#using webdriver. normally gets errors so I've found that using the Service() command let's it run smoothly!

##### Code #####

#read in list of stocks
df = pd.read_excel("stocks.xls")

#set up list with date
data = []
today = date.today()
data.append(today)

#set up service
service = Service("/Users/serena/Downloads/chromedriver")
service.start()

#for each stonk
for stock in df["stocks"]:
    
    # Use chrome webdriver
    driver = webdriver.Remote(service.service_url)
    
    # Get website
    yahoo = driver.get("https://au.finance.yahoo.com/")
    
    # Find search bar and enter a text
    search = driver.find_element(By.ID, 'yfin-usr-qry')
    search.send_keys("{0}".format(stock))
    time.sleep(5)
    pg = search.send_keys(Keys.RETURN)
    time.sleep(7)
    
    #extract previous closing value
    prev_close = driver.find_element(By.XPATH, "//*[@id='quote-summary']/div[1]/table/tbody/tr[1]/td[2]").get_attribute("innerHTML")
    
    #appending to list
    data.append(prev_close)
    driver.close()
 
driver.quit()

#writing to excel

#pre-existing excel book
name = 'prev_closing.xlsx'
wb = load_workbook(name)
page = wb.active

#appending new row with the list of data values:
page.append(data)

wb.save(filename=name)


