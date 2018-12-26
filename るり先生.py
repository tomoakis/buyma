import unittest, time, requests, webbrowser, bs4, datetime
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import csv, os, re, xlrd, openpyxl
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.alert import Alert


###　ログイン　###
#################
browser = webdriver.Chrome()
browser.get("https://www.buyma.com/my/buyerorders/")
email = browser.find_element_by_id('txtLoginId')
email.send_keys('taira4420@gmail.com')
password = browser.find_element_by_id('txtLoginPass')
password.send_keys('taira442054')
time.sleep(5)
browser.find_element_by_id('login_do').click()



