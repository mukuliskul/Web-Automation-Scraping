# Import necessary libraries
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import sys
import time
from docx import Document
import xlwings as xw
import math
import selenium

# Excel file containing the list of titles
exampleExcel = 'exampleExcelOfPublications'

# Load Excel workbook and specify the sheet
wb = xw.Book(exampleExcel)
sht = wb.sheets('Bibliogrphic details')

# Department name and difference value for cell count
deptname = 'zoology'
difference = int(387)

# Prompt user to enter the FILE NUMBER to start processing
filecount = int(input('Enter FILE NUMBER --> '))  # check the numbering once of the continuous word files
cellcount = filecount + difference
print(' ')

# Extract title list from Excel sheet
title_list = sht.range(f'C{cellcount}:C1000').value

# Initialize Selenium WebDriver (Chrome) and wait
driver = webdriver.Chrome('/Users/mukulsharma/Desktop/CS/python/automation/chromedriver')
wait = WebDriverWait(driver, 200)
driver.get('https://scholar.google.com')
time.sleep(0.8)
cellnumbercount = int(cellcount)

# Loop through each title in the list
for title in title_list:
    cellnumbercount += 1
    renamed = title[0:30]
    APAdocx = f'{filecount} {deptname} {renamed}.docx'
    APAlist = []
    count = 0

    # Search for the title on Google Scholar
    searchBox = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="gs_hdr_tsi"]')))
    searchBox.send_keys(title)
    time.sleep(1)
    search = webdriver.ActionChains(driver).send_keys(Keys.ENTER).perform()
    command = input("Enter Operation --> ")

    # Check if the user wants to skip processing this title
    if command == 'n' or command == 'N':
        filecount += 1
        driver.get('https://scholar.google.com')
        print(cellnumbercount)
        print(filecount)
        continue

    else:
        # Click on the "Cited by" link to get the list of citations
        citedbyLink = wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, 'Cited by')))
        pagetotaltxt = citedbyLink.text
        pagetotal = pagetotaltxt.lstrip('Cited by ')
        citedbyLink.click()

        # Get the list of "Cite" links for the current page and extract the citation details
        citeLINK = driver.find_elements(By.LINK_TEXT, 'Cite')
        for i in citeLINK:
            count += 1
            time.sleep(1)
            i.click()
            time.sleep(1)
            APAclick = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="gs_citt"]/table/tbody/tr[2]/td/div')))
            APAtext = str(count) + '. ' + str(APAclick.text)
            APAlist.append(APAtext)
            time.sleep(1)
            webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()

        # Calculate the number of pages with citations and navigate to each page
        totaldecimal = int(pagetotal) / 10
        totalroundoff = math.ceil(float(totaldecimal))
        finaltotal = int(totalroundoff) + 1
        finaltotal2 = int(totalroundoff) - 1

        if finaltotal <= 11:
            # If the number of pages is less than or equal to 11, navigate to each page and extract citations
            for i in range(2, int(finaltotal)):
                next = driver.find_element(By.LINK_TEXT, str(i)).click()
                citeLINK = []
                citeLINK = driver.find_elements(By.LINK_TEXT, 'Cite')
                for i in citeLINK:
                    count += 1
                    time.sleep(1.5)
                    i.click()
                    time.sleep(1)
                    APAclick = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="gs_citt"]/table/tbody/tr[2]/td/div')))
                    APAtext = str(count) + '. ' + str(APAclick.text)
                    APAlist.append(APAtext)
                    time.sleep(1)
                    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        else:
            # If there are more than 11 pages, navigate to the next pages and extract citations
            for i in range(int(finaltotal2)):
                next = driver.find_element(By.LINK_TEXT, 'Next').click()
                citeLINK = []
                citeLINK = driver.find_elements(By.LINK_TEXT, 'Cite')
                for i in citeLINK:
                    count += 1
                    time.sleep(1.5)
                    i.click()
                    time.sleep(1)
                    APAclick = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="gs_citt"]/table/tbody/tr[2]/td/div')))
                    APAtext = str(count) + '. ' + str(APAclick.text)
                    APAlist.append(APAtext)
                    time.sleep(1)
                    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()

        # Create a Word document and save the title and citation list
        document = Document()
        document.add_paragraph(title)
        for i in APAlist:
            para = document.add_paragraph(i)
        document.save(APAdocx)

        filecount += 1
        driver.get('https://scholar.google.com')

        print(cellnumbercount)
        print(filecount)
