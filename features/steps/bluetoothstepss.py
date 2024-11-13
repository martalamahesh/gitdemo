from behave import *
from selenium import webdriver
import time
from selenium import *
from openpyxl import Workbook

from selenium.webdriver.common.by import By
from openpyxl import *


@given(u':extract bluetooth data')
def step1(context):
    context.driver = webdriver.Chrome()
    context.driver.get("https://launchstudio.bluetooth.com/listings/search")
    adva = context.driver.find_element(By.XPATH, "//a[@id='advancedToggle']")
    adva.click()
    check = context.driver.find_element(By.XPATH, "//div[@id='searchIn']/label[3]/input")
    time.sleep(5)
    check.click()
    time.sleep(5)
    from_date = context.driver.find_element(By.XPATH,
                                            "//*[@id='advancedForm']/div/div[2]/div/div[5]/div[1]/div[2]/input")
    from_date.send_keys("2023-01-01")
    to_date = context.driver.find_element(By.XPATH,
                                          "//*[@id='advancedForm']/div/div[2]/div/div[6]/div/div[2]//input[@placeholder='YYYY-MM-DD']")
    to_date.send_keys("2023-01-03")
    search_box = context.driver.find_element(By.XPATH, "//*[@id='listingsSearch']/div/div[3]/div/div/div[1]/div/input")
    search_box.send_keys("BL")
    search = context.driver.find_element(By.XPATH, "//*[@id='searchButton']")
    search.click()
    time.sleep(10)
    wb = load_workbook('sample.xlsx')

    sheet = wb.active
    # the below code can be make dynamic also.it take some time
    for i in range(1, 11):
        for j in range(1, 7):
            a = "//*[@id='listingsSearch']/div/div[4]/table/tbody/tr[{tr_number}]/td[{td_number}]".format(
                tr_number=i, td_number=j
            )
            temp = context.driver.find_element(By.XPATH, a)
            sheet.cell(row=i, column=j).value = temp.text
    wb.save("sample.xlsx")
    wb.close()
    time.sleep(10)
