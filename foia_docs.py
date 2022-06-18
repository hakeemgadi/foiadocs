import os
from selenium import webdriver
import datetime
import time
import urllib.request
import openpyxl
import pickle

date_format = '%Y-%m-%d'
def get_date():
        while True:
                date=input("Please enter the date in the YYYY-MM-DD format\n")

                try:
                        date_obj = datetime.datetime.strptime(date, date_format)
                        break
                except ValueError:
                        print("Incorrect data format, should be YYYY-MM-DD. Please try again\n")
        return date_obj.strftime('%Y-%m-%d')

        
option = webdriver.ChromeOptions()
chrome_prefs = {}
option.experimental_options["prefs"] = chrome_prefs
#Disable image download
chrome_prefs["profile.default_content_settings"] = {"images": 2}
chrome_prefs["profile.managed_default_content_settings"] = {"images": 2}
driver = webdriver.Chrome(options=option,executable_path=os.path.join("..","..","chromedriver.exe"))

#Get input from user. Search string and date range
search_string = input('Please enter the search term:\n')
print('Please enter the search begin date:\n')
beginDate = get_date() #yyyymmdd
print('Please enter the search end date: \n')
endDate= get_date()

print(f"Downloading files for search term [{search_string}] beginning in {beginDate} and ending in {endDate}")

#Navigate to the application home page
driver.get(f"https://foia.state.gov/Search/Results.aspx?searchText={search_string}&beginDate={beginDate}&endDate={endDate}&publishedBeginDate=&publishedEndDate=&caseNumber=")

#Create workbook to store tabulated data
wb = openpyxl.Workbook()
ws = wb.active



time.sleep(2)
page_sz=20

startpg=1 #Normally 1. Unless not starting in the middle


#Get end page number by reading the number of pages
end_pg=int(driver.find_element_by_id("lblPagesTop").text)

#Loop pages
for i in range(startpg-1,end_pg):
        
        #Get result table in page
        result_table=driver.find_element_by_id("tblResults")

        #Get all rows in table
        rows=result_table.find_elements_by_tag_name("tr")

        #Loop result table rows
        for j,row in enumerate(rows):

                #dpoint will store the fields in the current row. It is serialized later to serve as a backup for Excel rows.
                dpoint=[]

                #Get all fields in row
                data=row.find_elements_by_tag_name("td")

                #Populate Excel workbook 
                if len(data)!=0:
                        #Write text into row i*page_sz+j
                        cell_row = i*page_sz+j
                        
                        print("Prcessing row no.: {cell_row}")

                        #Subject column
                        ws.cell(cell_row,1).value=data[1].text
                        dpoint.append(data[1].text)

                        #Get pdf link from subject column
                        pdflink=data[1].find_element_by_tag_name("span").get_attribute("title")

                        #Get file name only
                        file = pdflink.split("/")[-1]

                        #Construct local path of pdf
                        ws.cell(cell_row,2).hyperlink=os.path.join("..","DocStore","PDFS",file)
                        dpoint.append(os.path.join("..","DocStore","PDFS",file))
                        
                        ws.cell(cell_row,2).value="Local link"
                        dpoint.append("Local link")
                        
                        #Construct a global weblink from site-local link
                        ws.cell(cell_row,3).hyperlink="https://foia.state.gov/"+pdflink
                        dpoint.append("https://foia.state.gov/"+pdflink)

                        
                        ws.cell(cell_row,3).value="FOIA link"
                        dpoint.append("FOIA link")

                        ws.cell(cell_row,4).value=file
                        dpoint.append(file)

                        #Process the remaining of fields: Document Date, From, To, Posted Date, and Case Number
                        for rest in range(2,7):
                                #Offset by 3 columns in the Excel workbook
                                ws.cell(cell_row,rest+3).value=data[rest].text
                                dpoint.append(data[rest].text)

                        #Download the PDF
                        urllib.request.urlretrieve("https://foia.state.gov/"+pdflink, os.path.join("..","DocStore","PDFS",file))
                        print(file)

                        dfile=open(os.path.join("..","DocStore","PKL",str(cell_row)+".pkl"),"wb")
                        pickle.dump(dpoint,dfile)
                        dfile.close()
                        
                        time.sleep(2)
        if i < end_pg-1:
                #Go to next page if not last page

                inputElement = driver.find_element_by_id("txtPageTop")
                inputElement.clear()
                inputElement.send_keys(i+2)

                btn=driver.find_element_by_id("btnJumpTop")

                btn.click()
                print("Going to next page ...")

        time.sleep(3)

#Save workbook
wb.save(os.path.join("..","DocStore",'FilesLog.xlsx')) 

