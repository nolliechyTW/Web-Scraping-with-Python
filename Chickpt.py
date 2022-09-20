from bs4 import BeautifulSoup as BS 
import requests
import xlwings as xw

# I want to scrape first page
res = requests.get("https://www.chickpt.com.tw/cases?page="+str(1)) 
html = BS(res.text) 

# open a new, blank Excel worksheet
wb = xw.Book() 
sheet = wb.sheets ["top20pt"] # name sheet
sheet.range("A1").value = "job title"
sheet.range("B1").value = "employee name"
sheet.range("C1").value = "salary"
sheet.range("D1").value = "updated time"
sheet.range("E1").value = "url"

# fine the name of job
name_list = html.findAll("h2", {"class":"job-info-title ellipsis-job-name ellipsis"}) 
# find the name of employer
empolyer_list = html.findAll("p", {"class":"mobile-job-company ellipsis-mobile-job-company ellipsis display-control show-mobile"}) 
# find salary
salary_list = html.findAll("span",{"class":"place"})
# find updated tim
update_list = html.findAll("span",{"class":"date-time is-flex flex-align-center"})

# export results to worksheet
row = 2
for name in name_list:
    sheet.range(f"A{row}").value = name.text.strip() # use strip to remove spaces at the beginning and at the end of the string
    row = row + 1
    
row = 2
for employer in empolyer_list:
    sheet.range(f"B{row}").value = employer.text.strip() 
    row = row + 1       
        
row = 2
for salary in salary_list:
    sheet.range(f"C{row}").value = salary.text.strip() 
    row = row + 1

row = 2
for time in update_list:
    sheet.range(f"D{row}").value = time.text.strip() 
    row = row + 1
    
# get url and export it to worksheet
row=2
for k in html.findAll('a',{"class":"layout-width job-list-item is-flex flex-start flex-row flex-align-center is-tra"}):
    sheet.range(f"E{row}").value = (k['href'])
    row = row + 1
    
# save my worksheet 
wb.save("top20_Chickpt.xlsx")
