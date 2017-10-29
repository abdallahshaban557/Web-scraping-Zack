#CTRL+ALT+N to Run the python script
#from urllib.request import urlopen as uReq
#import urllib.parse
from bs4 import BeautifulSoup

import requests
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from openpyxl import Workbook
from openpyxl import load_workbook

#x = urllib.request.urlopen('https://pythonprogramming.net')

try:

    #Stop remember me from showing up
    chrome_options = Options()
    chrome_options.add_experimental_option('prefs', {
        'credentials_enable_service': False,
        'profile': {
            'password_manager_enabled': False
        }
    })
    #Open the chrome driver, and navigate to the page
    driver = webdriver.Chrome("chromedriver/chromedriver.exe",chrome_options=chrome_options)
    driver.get("https://www.zacks.com/stock/research/AAP/earnings-announcements")

    #Isolate the input fields, and insert data to form
    # USERNAME = driver.find_element_by_css_selector('#user_email')
    # USERNAME.send_keys('bbliss@sandiego.edu')
    # PASSWORD = driver.find_element_by_css_selector('#user_password')
    # PASSWORD.send_keys('fsu-stu712bb')
    # LOGIN = driver.find_element_by_xpath("//input[@value='Sign in']")
    Ticker_Symbol = driver.find_element_by_name('t')

    #start loop here
    Ticker_Symbol.send_keys('ticket')

    #Select the drop down row list and select 100 as the row number
    Drop_Down_List = Select(driver.find_element_by_xpath('//*[@id="earnings_announcements_earnings_table_length"]/label/select'))
    Drop_Down_List.select_by_visible_text('100')

    #Row number in the website EPS table
    EPS_Table_Row_Number = 2

  

  
    #Go through the whole table of EPS
    while True:
        try:
            Date_Of_EPS = driver.find_element_by_xpath('//*[@id="earnings_announcements_earnings_table_wrapper"]/div[3]/div[3]/div[2]/div/table/tbody/tr['+ str(EPS_Table_Row_Number) +']/td')
            #Just for validation
            print(Date_Of_EPS.text)
            EPS_Table_Row_Number += 1
        except Exception as e:
            break
    


    #Submit Data
    # LOGIN.click()

    #Create results sheet
    results_file = Workbook()

    #synopsis work sheet
    synopsis_worksheet = results_file.create_sheet("synopsis")

    #links work_sheet
    links_worksheet = results_file.create_sheet("Document Links")

    #Create labels for the sheets
    links_worksheet.cell(row=1, column=1).value = "Date"
    links_worksheet.cell(row=1, column=2).value = "Source"
    links_worksheet.cell(row=1, column=3).value = "Label"
    links_worksheet.cell(row=1, column=4).value = "URL"
    links_worksheet.cell(row=1, column=5).value = "Ticker Symbol"

    synopsis_worksheet.cell(row=1, column=1).value = "Synopsis"
    synopsis_worksheet.cell(row=1, column=2).value = "Ticker Symbol"
    #Access the original excel sheet
    original_list = load_workbook('source_file/list.xlsx')

    #Get the exact sheet
    scraping_sheet = original_list.worksheets[0]

    #count number of rows in sheet (Minus one because there is a title)
    row_count = scraping_sheet.max_row - 1

    #Current result sheet row - this is a counter to know where to start writing from the iteration in the second loop
    results_sheet_row = 2

    #Starting the iterations of the scraping list
    for i in range(2, row_count+2):
        #Set cell ID
        cell_id = 'A'+str(i)

        #Retrieve value of cell
        cell_value = scraping_sheet[cell_id].value

        driver.get("https://www.activistshorts.com/search/results?company%5Btickers%5D="+cell_value)
        time.sleep(3)

        #Get to the last required page
        # link_to_search_result = driver.find_element_by_xpath('//*[@id="container"]/div/div/div[2]/div/table/tbody/tr[1]/td[1]/a')
        # link_to_search_result = driver.find_element_by_xpath('//*[@id="container"]/div/div/div[2]/div/table/tbody/tr/td[1]/a')
        link_to_search_result = driver.find_element_by_css_selector(".tablesorter tbody tr td a")
        link_to_search_result.click()
        time.sleep(3)

        #Synopsis copy
        synopsis = driver.find_element_by_css_selector(".synopsis p")
        # synopsis_worksheet['A'+i].value = synopsis.text
        # synopsis_worksheet['B'+i].value = cell_value

        #Set Synopsis
        synopsis_worksheet.cell(row=i, column=1).value = synopsis.text

        #Add Ticker
        synopsis_worksheet.cell(row=i, column=2).value = cell_value

        #Bottom links counter
        documents_rows = driver.find_elements_by_css_selector(".tablesorter tbody tr")

        #Find number of rows in documents table, we add one so that the loop below accounts for the header location
        size_documents_rows = len(documents_rows)
        time.sleep(3)

        #loop through TD elements
        for j in range(1, size_documents_rows+1):
            #Find Tr element data-url attribute that contains the URL
            url_cell = driver.find_element_by_xpath("//div/div/div[4]/div/div/div[8]/div/table/tbody/tr[" + str(j) + "]").get_attribute("data-url")
            #Copy URL to sheet
            links_worksheet.cell(row=results_sheet_row, column=4).value = url_cell
            #Copy ticket symbol to URL
            links_worksheet.cell(row=results_sheet_row, column=5).value = cell_value
            for z in range(1, 4):
                    single_cell = driver.find_element_by_xpath("//div/div/div[4]/div/div/div[8]/div/table/tbody/tr[" + str(j) + "]/td[" + str(z) + "]")
                    links_worksheet.cell(row=results_sheet_row, column=z).value = single_cell.text
            results_sheet_row = results_sheet_row+1
            # save results sheet
            results_file.save("results_file/result.xlsx")

        #Mark cell as done
        scraping_sheet['B'+str(i)].value = "Done"
        original_list.save("source_file/log.xlsx")


    #Delete unnecessary sheet
    empty_sheet = results_file.get_sheet_by_name('Sheet')
    results_file.remove_sheet(empty_sheet)



    #close the chrome page
    driver.close()

except Exception as e:
    print(str(e))