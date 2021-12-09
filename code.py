##### Imports #####
from selenium import webdriver
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By


##### Comments #####

#path is where the chromedriver is on my local computer - so that needs to be changed for everyone
#https://chromedriver.chromium.org/getting-started
#if on mac and can't chromedriver: https://stackoverflow.com/questions/60362018/macos-catalinav-10-15-3-error-chromedriver-cannot-be-opened-because-the-de

#driver = webdriver.Chrome("/Users/serena/Downloads/chromedriver") 
#using webdriver. normally gets errors so I've found that using the Service() command let's it run smoothly! - i'll change to service later I swear lmaooo

##### Code #####

#setting up stonks
stocks = ["CBA.AX", "NAB.AX", "WBC.AX", "ANZ.AX"]

#for each stonk
for stock in stocks:

    # Use chrome webdriver
    driver = webdriver.Chrome("/Users/serena/Downloads/chromedriver")
    
    # Get website
    yahoo = driver.get("https://au.finance.yahoo.com/")
    
    # Find search bar and enter a text
    search = driver.find_element_by_id('yfin-usr-qry')
    search.send_keys("{0}".format(stock))
    time.sleep(5)
    pg = search.send_keys(Keys.RETURN)
    time.sleep(7)
    
    # Extract previous closing value
    prev_close = driver.find_element_by_xpath("//*[@id='quote-header-info']/div[3]/div[1]/div/fin-streamer[1]").get_attribute("innerHTML")
    
    #printing the number
    print(prev_close)

