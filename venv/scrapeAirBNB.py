from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import datetime
from datetime import date


driver = webdriver.Chrome("C:/Users/Don/Desktop/path folder/chromedriver_win32/chromedriver.exe")
driver.get("https://www.airbnb.com/login")
driver.maximize_window()

driver.find_element_by_xpath("//button[@data-testid='social-auth-button-email']").click()
driver.find_element_by_xpath("//input[@data-testid='login-signup-email']").send_keys("airbnb@vacationstation.com")
driver.find_element_by_xpath("//input[@data-testid='login-signup-password']").send_keys("Bike!0280")
driver.find_element_by_xpath("//button[@data-testid='signup-login-submit-btn']").click()
WebDriverWait(driver, 20).until(EC.presence_of_element_located(
    (By.XPATH, "//img[@src='https://a0.muscache.com/im/users/26122808/profile_pic/1421280273/original.jpg?aki_policy=profile_medium']")))
driver.get("https://www.airbnb.com/hosting/reservations/upcoming")
# driver.find_element_by_xpath("//div[@aria-label='Main navigation menu']").click()
WebDriverWait(driver, 30).until(EC.presence_of_element_located(
    (By.XPATH, "//button[@aria-label='Close']")))
driver.find_element_by_xpath("//button[@aria-label='Close']").click()
driver.find_element_by_xpath("//div[contains(text(), 'Filter')]").click()
todayDate = datetime.datetime.now()
todayDateString = str(todayDate.strftime('%x'))
# todayDateString = str(todayDate.strftime('%m%d%y'))

driver.find_element_by_xpath("//input[@aria-label='From']").send_keys(todayDateString)
driver.find_element_by_xpath("//div[contains(text(), 'Filter by dates')]").click()
# driver.implicitly_wait(5)
WebDriverWait(driver, 5).until(EC.presence_of_element_located(
    (By.XPATH, "//span[contains(text(), 'Apply')]/parent::button")))
driver.find_element_by_xpath("//span[contains(text(), 'Apply')]/parent::button").click()
exportButton = driver.find_element_by_xpath("//div[contains(text(), 'Export')]/parent::div/parent::button")
WebDriverWait(driver, 10).until(EC.presence_of_element_located(
    (By.XPATH, "//li[contains(@data-id, 'page-1')]")))
resPages = driver.find_elements_by_xpath("//li[contains(@data-id, 'page-')]")
lengthPages = len(resPages)
for x in range(lengthPages):
    print("in the loop")
    driver.find_elements_by_xpath("//li[contains(@data-id, 'page-')]")[x].click()
    # driver.implicitly_wait(5)
    print("clicked page")
    exportButton.click()
    WebDriverWait(driver, 15).until(EC.presence_of_element_located(
        (By.XPATH, "//div[contains(text(), 'Download CSV file')]/parent::div/parent::button")))
    driver.find_element_by_xpath("//div[contains(text(), 'Download CSV file')]/parent::div/parent::button").click()

    driver.find_element_by_xpath("//span[contains(text(), 'Download')]/parent::a").click()
    driver.implicitly_wait(3)
    driver.find_element_by_xpath("//*[name()='svg' and @aria-label='Close']").click()
    # WebDriverWait(driver, 5).until(EC.presence_of_element_located(
    #     (By.XPATH, "//span[contains(text(), 'Download')]/parent::a")))



# airbnb@vacationstation.com
# Bike!0280