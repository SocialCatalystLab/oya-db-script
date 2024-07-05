# Olaf de Rohan Willner
# ocd4@cornell.edu
# Version: July 5, 2024

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Set options
firefox_options = FirefoxOptions()
# firefox_options.add_argument("--headless")  # Run in headless mode if you don't need a GUI
service = FirefoxService(GeckoDriverManager().install())
driver = webdriver.Firefox(service=service, options=firefox_options)

# -- Start script --

# Store the current window handle
main_window = driver.current_window_handle

# Open the login page
driver.get('https://backupfs.provisocloud.com:255/default.aspx')

# enter username
driver.find_element_by_id("txtUserID").send_keys("505")
# enter password
driver.find_element_by_id("txtUserPassword").send_keys("123456")
# click login
driver.find_element_by_id("cmbLogin").click()

# TEMP FIX TO OVERRIDE FFOX WARNINGS
time.sleep(10)
print("start post-login")

# Switch to popup window
for handle in driver.window_handles:
    if handle != main_window: 
        popup = handle
        driver.switch_to.window(popup)

# click report menu
driver.find_element_by_id("TreeView15t0").click() 
# click loans report
driver.find_element_by_id("TreeView15t46").click()
# click loan transaction report
driver.find_element_by_id("TreeView15t75").click()
# click loan disbursement report
driver.find_element_by_id("TreeView15t76").click()
# click loan disbursement report [general]
driver.find_element_by_id("TreeView15t77").click()

time.sleep(5)

# Switch to new popup window
for handle in driver.window_handles:
    if handle != main_window: 
        popup = handle
        driver.switch_to.window(popup)

# FUTURE: set date range in date fields
# from date
driver.find_element_by_id("txtDf").clear()
driver.find_element_by_id("txtDf").send_keys("7/1/2024")
# to date
driver.find_element_by_id("txtDt").clear()
driver.find_element_by_id("txtDt").send_keys("7/3/2024")

# click show report
button1 = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "Button1"))
    )
button1.click()

# click download report to excel

# click download report in Excel
button2 = WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.ID, "Button2"))
    )
button2.click()

time.sleep(120)

driver.quit()

# -- End script --