import time
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select


def login(driver):
    driver.get("https://www.buyma.com/my/sell/new/")
    email = driver.find_element_by_id('txtLoginId')
    email.send_keys('namitaketomi123@gmail.com')
    password = driver.find_element_by_id('txtLoginPass')
    password.send_keys('seasider0093')
    driver.find_element_by_id('login_do').click()
    driver.implicitly_wait(40)
    time.sleep(1)
    driver.get("https://www.buyma.com/my/sell/new/")


def select(driver, id, option):
    # 1st click on the Select-control
    action = ActionChains(driver)
    element = driver.find_element_by_id(f'react-select-{id}--value')
    driver.execute_script("arguments[0].scrollIntoView();", element)
    action.move_to_element(element).click().perform()
    time.sleep(1)
    # 2nd click on the Select-menu-outer option
    element = driver.find_element_by_id(f'react-select-{id}--option-{option}')
    action = ActionChains(driver)
    action.move_to_element(element).click().perform()
    time.sleep(1)
