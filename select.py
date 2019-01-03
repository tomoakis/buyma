#!/usr/bin/python

import time
from enum import Enum
from lib import actions
from selenium import webdriver
import yaml


class Select(Enum):
    CATEGORY = 2
    SUB_CATEGORY = 12
    ITEM = 13


category = 'レディースファッション'
subCategory = '財布・小物'
item = 'キーケース'

with open('./constants/カテゴリ.yml', 'r') as file:
    categoriesDoc = yaml.load(file)

with open(f'./constants/{category}.yml', 'r') as file:
    subCategoryDoc = yaml.load(file)

id = categoriesDoc["categories"][category]['id']
subId = categoriesDoc["categories"][category]['subcategories'][subCategory]
itemId = subCategoryDoc[subCategory][item]

driver = webdriver.Chrome(executable_path='./drivers/chromedriver_linux')
actions.login(driver)
actions.select(driver, Select.CATEGORY.value, id)
actions.select(driver, Select.SUB_CATEGORY.value, subId)
actions.select(driver, Select.ITEM.value, itemId)

time.sleep(3)
driver.close()
