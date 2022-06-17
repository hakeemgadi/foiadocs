import os
from selenium import webdriver
import datetime
import time
import pickle
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import urllib.request
import openpyxl
import pickle

date_format = '%Y-%m-%d'
def get_date():
        while True:
                date=input("Please enter the date in the YYYY-MM-DD format\n")

                try:
                        date_obj = datetime.datetime.strptime(date, date_format)
                        print(date_obj)
                        break
                except ValueError:
                        print("Incorrect data format, should be YYYY-MM-DD. Please try again\n")
        return date_obj.strftime('%Y%m%d')

        
option = webdriver.ChromeOptions()
chrome_prefs = {}
option.experimental_options["prefs"] = chrome_prefs
chrome_prefs["profile.default_content_settings"] = {"images": 2}
chrome_prefs["profile.managed_default_content_settings"] = {"images": 2}
driver = webdriver.Chrome(options=option,executable_path=os.path.join("..","..","chromedriver.exe"))
#Navigate to the application home page
print("anaosje;kf")

search_string = input('Please enter the search term\n')
print('Please enter the search begin date\n')
beginDate = get_date() #yyyymmdd
print('Please enter the search end date\n')
endDate= get_date()

print(f"Downloading files for search term [{search_string}] beginning in {beginDate} and ending in {endDate}")
driver.get(f"https://foia.state.gov/Search/Results.aspx?searchText={search_string}&beginDate={beginDate}&endDate={endDate}&publishedBeginDate=&publishedEndDate=&caseNumber=")

#Create workbook to store tabulated data
wb = openpyxl.Workbook()
ws = wb.active



time.sleep(2)
page_sz=20

startpg=1

end_pg=int(driver.find_element_by_id("lblPagesTop").text)

for i in range(startpg-1,end_pg):
        result_table=driver.find_element_by_id("tblResults")

        rows=result_table.find_elements_by_tag_name("tr")
        for j,row in enumerate(rows):

                dpoint=[]
                
                data=row.find_elements_by_tag_name("td")
                if len(data)!=0:
                        ws.cell(i*page_sz+j,1).value=data[1].text
                        dpoint.append(data[1].text)
                        
                        print(len(data),i*page_sz+j)
                        pdflink=data[1].find_element_by_tag_name("span").get_attribute("title")
                        
                        file = pdflink.split("/")[-1]

                        ws.cell(i*page_sz+j,2).hyperlink=os.path.join("..","DocStore","PDFS",file)
                        dpoint.append(os.path.join("..","DocStore","PDFS",file))
                        
                        ws.cell(i*page_sz+j,2).value="Local link"
                        dpoint.append("Local link")
                        

                        ws.cell(i*page_sz+j,3).hyperlink="https://foia.state.gov/"+pdflink
                        dpoint.append("https://foia.state.gov/"+pdflink)

                        
                        ws.cell(i*page_sz+j,3).value="FOIA link"
                        dpoint.append("FOIA link")

                        ws.cell(i*page_sz+j,4).value=file
                        dpoint.append(file)

                        
                        for rest in range(2,7):
                                ws.cell(i*page_sz+j,rest+3).value=data[rest].text
                                dpoint.append(data[rest].text)

                        urllib.request.urlretrieve("https://foia.state.gov/"+pdflink, os.path.join("..","DocStore","PDFS",file))
                        print(file)

                        dfile=open(os.path.join("..","DocStore","PKL",str(i*page_sz+j)+".pkl"),"wb")
                        pickle.dump(dpoint,dfile)
                        dfile.close()
                        
                        time.sleep(2)


        inputElement = driver.find_element_by_id("txtPageTop")
        inputElement.clear()
        inputElement.send_keys(i+2)

        btn=driver.find_element_by_id("btnJumpTop")

        btn.click()
        print("Click!")

        time.sleep(3)
wb.save(os.path.join("..","DocStore",'FilesLog.xlsx')) 

