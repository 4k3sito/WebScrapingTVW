"""
I am looking for script for the following website site -

https://www.wvsao.gov/CountyCollections/Default#STIInquiry

Tax Year - 2020
County - Should pull up each county (there are 55)
District - Must be set to Always equals ALL DISTRICTS

and scrape the following information from the resulting pages into a CSV file:

Cert Number / Ticket /Delinquent Name / Buyer Name / Description /Redeemed Date / NTR REC / Deed Fee Rec

The final product should include a python script in .py format.
"""
import time
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from openpyxl import Workbook

# URL and Excel File Creation
url = "https://www.wvsao.gov/CountyCollections/Default#STIInquiry"
excelFile = "Tax_Information_WV.xlsx"
file = Workbook()
fileActive = file.active
file.save("Tax_Information_WV.xlsx")
file.active
#Init the driver in Chrome
options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)
driver = webdriver.Chrome(options=options)
driver.get(url)
#Function to get the table from the current page
def pageTable():
    #Get page source 
    pageSource = driver.page_source
    #Create a HTML File for the page source
    fileToWrite = open("page_source.html", "w")
    fileToWrite.write(pageSource)
    fileToWrite.close()
    #Open page Source in BeautidulSoup
    with open("page_source.html", "r") as tableContent:
        soup = BeautifulSoup(tableContent, 'lxml')
    #Find Table in BS
    table = soup.find('table', {'class': 'table table-striped'})

    #Get the Headers
    headers = table.find_all("th")
    th = []
    for i in headers:
        headersName = i.text
        th.append(headersName)

    #Get the values of the Table
    tableValues = []
    for x in table.find_all('tr')[1:]:
        td_tags = x.find_all('td')
        td_vals = [y.text for y in td_tags]
        tableValues.append(td_vals)

    #Create a DataFrame and Export to XLSX
    df = pd.DataFrame(tableValues, columns=th)
    driver.back()
    time.sleep(2)

    return df


def country_List():
    dropdownSelectTag = driver.find_element(By.ID, "UltraWideContent_ddlSTICounties")
    countriesDroprown = Select(dropdownSelectTag)
    countryCounter = countriesDroprown.options
    countryList = [option.text for option in countryCounter]
    
    return countryList


def pageSelectTag():
    dropdownSelectTag = driver.find_element(By.ID, "UltraWideContent_ddlSTICounties")
    countriesDroprown = Select(dropdownSelectTag)

    return countriesDroprown

def menueSelection():
    countryList = country_List()
    with pd.ExcelWriter(excelFile, mode='a') as writer:
        for index in range(1, len(countryList) - 1):    
            countriesDroprown = pageSelectTag()
            countriesDroprown.select_by_index(index)
            countryName = countryList[index]

            button = driver.find_element(By.ID, "UltraWideContent_btnSTIResults")
            button.click()
            time.sleep(2)
            table = pageTable()

            table.to_excel(excel_writer= writer, sheet_name= countryName, index=False)

            

menueSelection()
file.save("Tax_Information_WV.xlsx")


