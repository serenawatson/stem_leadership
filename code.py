##### Imports #####
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

##### Comments #####

#path is where the chromedriver is on my local computer - so that needs to be changed for everyone
#https://chromedriver.chromium.org/getting-started
#if on mac and can't chromedriver: https://stackoverflow.com/questions/60362018/macos-catalinav-10-15-3-error-chromedriver-cannot-be-opened-because-the-de

#driver = webdriver.Chrome("/Users/serena/Downloads/chromedriver") 
#using webdriver. normally gets errors so I've found that using the Service() command let's it run smoothly!

##### Code #####

# Use chrome webdriver
s = Service("C:\Program Files (x86)\Chrome Driver\chromedriver.exe")
driver = webdriver.Chrome(service= s)

# Get website
driver.get("https://au.finance.yahoo.com/")

# Find search bar and enter a text
search = driver.find_element(By.NAME, "yfin-usr-qry")
search.send_keys("CBA.AX")
search.send_keys(Keys.RETURN)

# Program waits 10 seconds before finding the element
try:
    main = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "quote-summary"))
    )
    print(main.text)
finally:
    time.sleep(3)
    driver.quit()

