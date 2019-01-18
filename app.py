import requests
from bs4 import BeautifulSoup
import csv
from pandas.io.excel import ExcelWriter
import pandas

# base url
URL = "https://columbusrealtors.com/find.aspx?mode=browse&letter="
# define session
session = requests.session()
# define header of request
header = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/67.0.3396.99 Safari/537.36'}
# define file names
csv_file = "result.csv"
xls_file = "result.xlsx"
# define the pageNumber range when page size is 35
maxPagenumber = 235
cnt = 0


# define functions

# getValue
## param soup : beautiful soup Object:
## param eleName :  element Name
## param dict   : query for element
## return mixed
def getValue(soup, eleName, dict):
    try:
        ele = soup.find(eleName, dict)
        return ele.get('value')
    except:
        return ""
        pass
# savetoCSV
def writeToFile(data, isHeader=False):
    if isHeader:
        myFile = open(csv_file, 'w', newline='')
    else:
        myFile = open(csv_file, 'a', newline='')
    with myFile:
        writer = csv.writer(myFile)
        writer.writerow(data)
    myFile.close()

# step 1
# get html content for get hidden values
res = session.get(URL)
soup = BeautifulSoup(res.text, "html.parser")
__VIEWSTATE = getValue(soup, "input", {'id':"__VIEWSTATE"})
__VIEWSTATEGENERATOR = getValue(soup, "input", {"id" : "__VIEWSTATEGENERATOR"})
__EVENTVALIDATION = getValue(soup, "input", {"id" : "__EVENTVALIDATION"})


def getPageData(pageNumber):
    global __VIEWSTATE, __VIEWSTATEGENERATOR, __EVENTVALIDATION, header, cnt
    if pageNumber < 3:
        __EVENTTARGET = "ctl00$body$primary_body_1$ctl01$ucSearchResults$radSearchResults$ctl00$ctl02$ctl00$ctl0" + str(5 + pageNumber * 2)
    else:
        __EVENTTARGET = "ctl00$body$primary_body_1$ctl01$ucSearchResults$radSearchResults$ctl00$ctl02$ctl00$ctl" + str(5  + pageNumber* 2)
    data = {
        "__EVENTTARGET": __EVENTTARGET,
        "__VIEWSTATE": __VIEWSTATE,
        "__VIEWSTATEGENERATOR": __VIEWSTATEGENERATOR,
        "__EVENTVALIDATION": __EVENTVALIDATION,
    }

    # get page data of list
    res = session.post(URL, data=data,headers=header)
    soup1 = BeautifulSoup(res.text, "html.parser")
    __VIEWSTATE = getValue(soup1, "input", {'id': "__VIEWSTATE"})
    __VIEWSTATEGENERATOR = getValue(soup1, "input", {"id": "__VIEWSTATEGENERATOR"})
    __EVENTVALIDATION = getValue(soup1, "input", {"id": "__EVENTVALIDATION"})

    data = {
        "__EVENTTARGET": "",
        "__VIEWSTATE": __VIEWSTATE,
        "__VIEWSTATEGENERATOR": __VIEWSTATEGENERATOR,
        "__EVENTVALIDATION": __EVENTVALIDATION,
    }

    dataTable = soup1.find("table", {"id" : "ctl00_body_primary_body_1_ctl01_ucSearchResults_radSearchResults_ctl00"}).find_all("tbody")[2]
    for tr in dataTable.find_all('tr'):
        # first name, last name, company
        lastName = tr.find_all("td")[1].text
        firstName= tr.find_all("td")[2].text
        company = tr.find_all("td")[3].text
        city = tr.find_all("td")[4].text
        type = tr.find_all("td")[5].text.replace('\r\n\t\t\t\t\t\t\t\t','').replace('\r\n\t\t\t\t\t\t\t','')


        # get other data (address, Phone, Fax, Email Address) from detail page
        href = tr.find('a').get('href')
        data.update({"__EVENTTARGET": href.replace("javascript:__doPostBack('", "").replace("','')", "")})
        # get detail data
        res1 = session.post(URL, data=data, headers=header)
        soup2 = BeautifulSoup(res1.text, "html.parser")
        dataDiv = soup2.find_all("div", {"class": "island"})[1]
        ptags = dataDiv.find_all("p")
        #address
        try:
            text = ptags[0].text
            address = text.replace('\r\n\t\t\t\t','').replace(company, '').replace('\r\n\t\t\t','').replace('\n','')
        except:
            address = ""
            pass
        #phone
        phone = ptags[1].text.replace('Phone:','').replace('\n ','')
        #fax
        fax = ptags[2].text.replace('Fax:','').replace('\n ','')
        #email
        email = ptags[3].text.replace('Email:','').replace('\n ','')
        dt = [lastName, firstName, company, city, type, address, phone, fax, email]
        cnt += 1
        print(cnt)
        # save one data to csv
        writeToFile(dt)

csvHeader  = ["Last Name", "First Name", "Company", "City", "Type", "Address", "Phone", "Fax", "Email Address"]
writeToFile(csvHeader, True)

pageNumber = 0
isFirst = True
for number in range(0, maxPagenumber):
    if isFirst == True and pageNumber <= 10:
        pageNumber = pageNumber
    elif isFirst == False and pageNumber <= 11:
        pageNumber = pageNumber
    else:
        isFirst = False
        pageNumber = pageNumber % 10
    print("pageNumber" + str(pageNumber))
    getPageData(pageNumber)
    pageNumber += 1

# convert csv file to excel format
with ExcelWriter(xls_file) as ew:
    df = pandas.read_csv(csv_file)
    df.to_excel(ew, sheet_name="sheet1", index=False)
