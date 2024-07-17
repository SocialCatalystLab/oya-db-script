# Olaf de Rohan Willner
# ocd4@cornell.edu / olaf.willner@econ.uzh.ch
# Version: July 5, 2024

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from datetime import datetime
import shutil
import os
import time

# Set options
firefox_options = FirefoxOptions()
# firefox_options.add_argument("--headless")  # Run in headless mode if you don't need a GUI
service = FirefoxService(GeckoDriverManager().install())
driver = webdriver.Firefox(service=service, options=firefox_options)

# Files to fetch:
# Loan status report DONE
# Loan assessment report DONE
# Loan disbursement report DONE
# Loan applications report DONE

# -- Start LOAN STATUS script --

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
# click loan status report
driver.find_element_by_id("TreeView15t120").click()
# click loan status report
driver.find_element_by_id("TreeView15t121").click()

time.sleep(5)

# Switch to new popup window
for handle in driver.window_handles:
    if handle != main_window: 
        popup = handle
        driver.switch_to.window(popup)

# click download report in Excel
button3 = WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.ID, "Button3"))
    )
button3.click()

# accept alert dialog
driver.switch_to.alert.accept()

time.sleep(120)

# -- End LOAN STATUS script --

# -- Start LOAN APPLICATIONS script --

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
# click loan application report
driver.find_element_by_id("TreeView15t47").click()
# click loan application report
driver.find_element_by_id("TreeView15t48").click()

time.sleep(5)

# Switch to new popup window
for handle in driver.window_handles:
    if handle != main_window: 
        popup = handle
        driver.switch_to.window(popup)

# DEFAULT TO TODAY'S DATE
# from date
driver.find_element_by_id("txtDf").clear()
driver.find_element_by_id("txtDf").send_keys("1/1/2024")
# to date
driver.find_element_by_id("txtDt").clear()
driver.find_element_by_id("txtDt").send_keys(datetime.today().strftime('%m/%d/%Y'))

# click show report
button1 = WebDriverWait(driver, 600).until(
        EC.presence_of_element_located((By.ID, "Button1"))
    )
button1.click()

# click download report in Excel
button2 = WebDriverWait(driver, 600).until(
        EC.presence_of_element_located((By.ID, "Button2"))
    )
button2.click()

time.sleep(120)

# -- End LOAN APPLICATIONS script --

# -- Start LOAN REPAYMENT script --

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
# click loan repayment report
driver.find_element_by_id("TreeView15t81").click()
# click received
driver.find_element_by_id("TreeView15t86").click()
# click loan repayment report [RECEIVED]
driver.find_element_by_id("TreeView15t87").click()

time.sleep(5)

# Switch to new popup window
for handle in driver.window_handles:
    if handle != main_window: 
        popup = handle
        driver.switch_to.window(popup)

# DEFAULT TO TODAY'S DATE
# from date
driver.find_element_by_id("txtDf").clear()
driver.find_element_by_id("txtDf").send_keys("7/1/2024")
# to date
driver.find_element_by_id("txtDt").clear()
driver.find_element_by_id("txtDt").send_keys(datetime.today().strftime('%m/%d/%Y'))

# click download report in Excel
button2 = WebDriverWait(driver, 600).until(
        EC.presence_of_element_located((By.ID, "Button2"))
    )
button2.click()

time.sleep(120)

os.rename("path/to/current/file.foo", "path/to/new/destination/for/file.foo")
os.replace("path/to/current/file.foo", "path/to/new/destination/for/file.foo")
shutil.move("path/to/current/file.foo", "path/to/new/destination/for/file.foo")

# -- End LOAN REPAYMENT script --


# -- Start LOAN ASSESSMENT script --

# # Store the current window handle
# main_window = driver.current_window_handle

# # Open the login page
# driver.get('https://backupfs.provisocloud.com:255/default.aspx')

# # enter username
# driver.find_element_by_id("txtUserID").send_keys("505")
# # enter password
# driver.find_element_by_id("txtUserPassword").send_keys("123456")
# # click login
# driver.find_element_by_id("cmbLogin").click()

# # TEMP FIX TO OVERRIDE FFOX WARNINGS
# time.sleep(10)
# print("start post-login")

# # Switch to popup window
# for handle in driver.window_handles:
#     if handle != main_window: 
#         popup = handle
#         driver.switch_to.window(popup)

# # click report menu
# driver.find_element_by_id("TreeView15t0").click() 
# # click loans report
# driver.find_element_by_id("TreeView15t46").click()
# # click loan application report
# driver.find_element_by_id("TreeView15t47").click()
# # click loan assessment report
# driver.find_element_by_id("TreeView15t56").click()
# # click loan assessment report [general]
# driver.find_element_by_id("TreeView15t57").click()

# time.sleep(5)

# # Switch to new popup window
# for handle in driver.window_handles:
#     if handle != main_window: 
#         popup = handle
#         driver.switch_to.window(popup)

# # OPTIONAL: set date range in date fields
# # from date
# # driver.find_element_by_id("txtDf").clear()
# # driver.find_element_by_id("txtDf").send_keys("7/1/2024")
# # to date
# # driver.find_element_by_id("txtDt").clear()
# # driver.find_element_by_id("txtDt").send_keys("7/3/2024")

# # click show report
# button1 = WebDriverWait(driver, 20).until(
#         EC.presence_of_element_located((By.NAME, "Button1"))
#     )
# button1.click()

# # click export
# exportbutton = WebDriverWait(driver, 20).until(
#         EC.presence_of_element_located((By.NAME, "CrystalReportViewer1$ctl02$ctl00"))
#     )
# exportbutton.click() # WORKS UNTIL HERE

# # Switch to new popup window
# for handle in driver.window_handles:
#     if handle != main_window: 
#         popup = handle
#         driver.switch_to.window(popup)

# time.sleep(5)

# # select MS Excel 97-2000 (Data Only)
# select = Select(driver.find_element_by_name('exportformat'))
# select.select_by_value('RecordToMSExcel')

# # click ok
# driver.find_element_by_id("submitexport").click()

# time.sleep(10)

# driver.quit()

# # -- End LOAN ASSESSMENT script --

# # -- Start LOAN DISBURSEMENT script --

# # initialize new driver
# firefox_options = FirefoxOptions()
# service = FirefoxService(GeckoDriverManager().install())
# driver = webdriver.Firefox(service=service, options=firefox_options)

# # Store the current window handle
# main_window = driver.current_window_handle

# # Open the login page
# driver.get('https://backupfs.provisocloud.com:255/default.aspx')

# # enter username
# driver.find_element_by_id("txtUserID").send_keys("505")
# # enter password
# driver.find_element_by_id("txtUserPassword").send_keys("123456")
# # click login
# driver.find_element_by_id("cmbLogin").click()

# # TEMP FIX TO OVERRIDE FFOX WARNINGS
# time.sleep(10)
# print("start post-login")

# # Switch to popup window
# for handle in driver.window_handles:
#     if handle != main_window: 
#         popup = handle
#         driver.switch_to.window(popup)

# # click report menu
# driver.find_element_by_id("TreeView15t0").click() 
# # click loans report
# driver.find_element_by_id("TreeView15t46").click()
# # click loan transaction report
# driver.find_element_by_id("TreeView15t75").click()
# # click loan disbursement report
# driver.find_element_by_id("TreeView15t76").click()
# # click loan disbursement report [general]
# driver.find_element_by_id("TreeView15t77").click()

# time.sleep(5)

# # Switch to new popup window
# for handle in driver.window_handles:
#     if handle != main_window: 
#         popup = handle
#         driver.switch_to.window(popup)

# # FUTURE: set date range in date fields
# # from date
# driver.find_element_by_id("txtDf").clear()
# driver.find_element_by_id("txtDf").send_keys("7/1/2024")
# # to date
# driver.find_element_by_id("txtDt").clear()
# driver.find_element_by_id("txtDt").send_keys("7/3/2024")

# # click show report
# button1 = WebDriverWait(driver, 20).until(
#         EC.presence_of_element_located((By.ID, "Button1"))
#     )
# button1.click()

# # click download report in Excel
# button2 = WebDriverWait(driver, 60).until(
#         EC.presence_of_element_located((By.ID, "Button2"))
#     )
# button2.click()

# time.sleep(120)


# # -- End LOAN DISBURSEMENT script --

# driver.quit()