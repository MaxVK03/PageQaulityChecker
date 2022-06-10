# Imports are below here
import numpy as np
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import re
import requests
from collections import Counter
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
# Below are variables that are system dependent
#Change these on your instance of the program
#The below is the Location of the Excel sheet that reads in the URL's
#Make sure that they have https:// in front of the url
ExcelLocation = r"LinkCatExcel.xlsx"
#Below is the location of the webdriver
#In this program we use chromedriver, can be found https://chromedriver.chromium.org/downloads
# webDriverLoc = r"C:\Users\maxvo\Desktop\PythonInOut\WebDriver\chromedriver.exe"
#Below is the location of the excel we print out to, make sure it is not open while running]
#The program otherwise it will fail
ExOut = r"CheckOut.xlsx"
#Below is where the text file that contains a list of words that you want to classify as bad words.
# The program looks for exact match of the words and if found sets bad words to true.
BadWordsText =  r"TextIn.txt"
# driver = webdriver.Chrome(webDriverLoc)


# Gets the links from excel and converts it to a list for python
def readExcel(FileLoc):
    exIn = pd.read_excel(FileLoc)
    links = list(exIn["URL"].values)
    return links

#This function uses the selenium driver to get the title of the webpages
def driverGetTitles(driver):
    tit = driver.title
    return tit

#This function uses the selenium driver to get all body text of a page
def driverGetPageText(driver):
    try:
        driver.set_page_load_timeout(10)
        texttt = ""
        # driver.manage().timeouts().pageLoadTimeout(15, TimeUnit.SECONDS)
        # al = driver.find_elements_by_css_selector('p')
        al = driver.find_element(By.XPATH, "/html/body").text
        texttt = al
        return texttt
    except:
        return "None"

def ReqGetHeaders(link):
    try:
        response = requests.head(link)
        return response.status_code
    except:
        return 404


def driverGetBadWords(driver):
    try:
        NotNice = False
        texttt = ""
        driver.set_page_load_timeout(5)
        al = driver.find_element(By.TAG_NAME, "body")
        texttt = al.text

        with open(BadWordsText) as f:
            lines = f.readlines()

        badWords = []
        for i in lines:
            if i != "\n":
                badWords.append(i.replace("\n",""))

        for i in badWords:
            if (" " + i + " ") in texttt:
                NotNice = True

        return str(NotNice)

    except:
        return "None"

def getWordCount(text):
    text.replace("  "," ")
    text.replace("  ", " ")
    text.replace("  ", " ")
    text.replace("\n", " ")
    return str(text.count(" "))

def getSemRush(link):
    SemLink = "https://api.semrush.com/?key=d54660a9bf2e19552ff7bc3b5b64e4d9&type=domain_ranks&export_columns=Dn,Rk,Or,Ot,Oc&domain=INSDOM!@&database=us"
    Outlink = SemLink.replace("INSDOM!@", link)
    response = requests.get(Outlink).text
    Domain = ""
    Rank = ""
    OrgKW = ""
    OrgTraf = ""
    OrgCost = ""
    semRes = response.splitlines()
    try:
        ress = semRes[1].split(";")
        Domain = str(ress[0])
        Rank = str(ress[1])
        OrgKW = str(ress[2])
        OrgTraf = str(ress[3])
        OrgCost = str(ress[4])

    except:
        Domain = "Nan"
        Rank = "Nan"
        OrgKW = "Nan"
        OrgTraf = "Nan"
        OrgCost = "Nan"

    return [link ,Domain, Rank, OrgKW, OrgTraf, OrgCost]

def getTop5KW(driver):
    try:
        StopWords = []
        al = driver.find_element(By.TAG_NAME, "body")
        text = al.text
        # Below is our stop words. Feel free to add where needed
        newText = text.lower()
        newText = newText.replace(" the "," ")
        newText = newText.replace(" and ", " ")
        newText = newText.replace(" in ", " ")
        newText = newText.replace(" how ", " ")
        newText = newText.replace(" or ", " ")
        newText = newText.replace(" a ", " ")
        newText = newText.replace(" to ", " ")
        newText = newText.replace(" of ", " ")
        newText = newText.replace(" we ", " ")
        newText = newText.replace(" i ", " ")
        newText = newText.replace(" you ", " ")
        newText = newText.replace(" us ", " ")
        newText = newText.replace(" is ", " ")
        newText = newText.replace(" our ", " ")
        newText = newText.replace(" re ", " ")
        newText = newText.replace(" i ", " ")
        newText = newText.replace(" you ", " ")
        newText = newText.replace(" us ", " ")
        newText = newText.replace(" are ", " ")
        newText = newText.replace(" thia ", " ")
        newText = newText.replace(" as ", " ")
        newText = newText.replace(" what ", " ")
        newText = newText.replace(" own ", " ")
        newText = newText.replace(" this ", " ")
        newText = newText.replace(" will ", " ")
        newText = newText.replace(" ever ", " ")
        newText = newText.replace(" while ", " ")
        newText = newText.replace(" data ", " ")
        newText = newText.replace(" on ", " ")
        newText = newText.replace(" your ", " ")
        newText = newText.replace(" that ", " ")
        newText = newText.replace(" ever ", " ")
        newText = newText.replace(" for ", " ")
        newText = newText.replace(" with ", " ")
        newText = newText.replace(" by ", " ")
        newText = newText.replace(" it ", " ")
        newText = newText.replace(" up ", " ")
        newText = newText.replace(" · ", " ")
        newText = newText.replace(" an ", " ")
        newText = newText.replace(" ago ", " ")
        data_set = newText
        split_it = data_set.split()
        # Pass the split_it list to instance of Counter class.
        Counters_found = Counter(split_it)
        # print(Counters)
        # most_common() produces k frequently encountered
        # input values and their respective counts.
        most_occur = Counters_found.most_common(5)
        # print(most_occur)

    except:
        most_occur = ["none", "none","none","none","none"]
    return most_occur




# Above here is all methods
ListOfLinks = readExcel(ExcelLocation)
LinksNotWorking = []
print(ListOfLinks)
LinksRedirect = []
LinksAndServerHeaders = {}

for i in ListOfLinks:
    header = ReqGetHeaders(i)
    LinksAndServerHeaders[i] = header

LinksAndContainsBadWords = {}
LinksAndText = {}
LinksAndTitles = {}
LinkAndWordCount = {}
LinksAndTop5kw = {}
LinksAndSem = {}


for i in ListOfLinks:
    try:
        if ((str(LinksAndServerHeaders.get(i))[0:2]) != "40"):
            LinksAndSem[i] = getSemRush(i)
            options = Options()
            options.headless = True
            service = Service('chromedriver.exe')
            driver = webdriver.Chrome(service=service, options=options)
            driver.set_page_load_timeout(15)
            print("Getting " + i)
            driver.get(i)
            LinksAndText[i] = driverGetPageText(driver)
            LinksAndContainsBadWords[i] = driverGetBadWords(driver)
            LinksAndTitles[i] = driverGetTitles(driver)
            LinkAndWordCount[i] = getWordCount(LinksAndText.get(i))
            LinksAndTop5kw[i] = getTop5KW(driver)
            driver.close()

        else:
            print("Links not working/ Server Code 40_: "  + i)
            LinkAndWordCount[i] = "None"
            LinksAndSem[i] = ["None", "None", "None", "None", "None", "None"]
            LinksAndContainsBadWords[i] = "None"
            LinksAndText[i] = "None"
            LinksAndTitles[i] = "None"
            LinksAndTop5kw[i] = "Nonee"

    except:
        print("Obtaining Link not working: " + i)
        LinkAndWordCount[i] = "TimeOut"
        # LinksAndSem[i] = ["TimeOut", "TimeOut", "TimeOut", "TimeOut", "TimeOut", "TimeOut"]
        LinksAndContainsBadWords[i] = "TimeOut"
        LinksAndText[i] = "TimeOut"
        LinksAndTitles[i] = "TimeOut"
        LinksAndTop5kw[i] = "TimeOut"

ArrTitles = []
ArrWordCount = []
ArrBadWords = []
ArrTop5KW = []
ArrServerHeads = []
ArrDomain = []

for i in LinksAndServerHeaders.keys():
    # print(i + ": " + str(LinksAndServerHeaders.get(i)))
    ArrServerHeads.append(str(LinksAndServerHeaders.get(i)))

for i in ListOfLinks:
    try:
        # print(i + ": " + str(LinksAndTitles.get(i)))
        ArrTitles.append(str(LinksAndTitles.get(i)))

    except:
        ArrTitles.append("Nan")

for i in ListOfLinks:
    try:
        # print(i + ": " + str(LinksAndTop5kw.get(i)))
        ArrTop5KW.append(str(LinksAndTop5kw.get(i)).replace("[","").replace("]",""))
    except:
        ArrTop5KW.append("Nan")

# print("Contains bad words")
for i in LinksAndContainsBadWords:
    try:
        # print(i + ": " + str(LinksAndContainsBadWords.get(i)))
        ArrBadWords.append(str(LinksAndContainsBadWords.get(i)))
    except:
        ArrBadWords.append("Nan")

for i in LinkAndWordCount:

    if i in LinkAndWordCount.keys():
        # print(i  + ": " + str(LinkAndWordCount.get(i)))
        ArrWordCount.append(LinkAndWordCount.get(i))


# print("-==--=-==-=--=-=")
# print(LinksAndSem)

Domains = []
semRanks = []
OrgKWs = []
OrgTrafs = []
OrgCosts = []


for i in LinksAndSem.keys():
    SemProf = LinksAndSem.get(i)
    Domains.append(SemProf[1])
    semRanks.append(SemProf[2])
    OrgKWs.append(SemProf[3])
    OrgTrafs.append(SemProf[4])
    OrgCosts.append(SemProf[5])



CheckedPage = pd.DataFrame()


CheckedPage = pd.DataFrame()
toAdd1 = np.array((ListOfLinks))
toAdd2 = np.array(ArrTitles)
toAdd3 = np.array(ArrTop5KW)
toAdd4 = np.array(ArrServerHeads)
toAdd5 = np.array(ArrBadWords)
toAdd6 = np.array(ArrWordCount)
toAdd7 = np.array(semRanks)
toAdd8 = np.array(OrgKWs)
toAdd9 = np.array(OrgTrafs)
toAdd10 = np.array(OrgCosts)
toAdd11 = np.array(ArrBadWords)
try:
    CheckedPage["URL"] = toAdd1
    print("URL Sucsess")
except:
    print("URL Failed")

try:
    CheckedPage["Title"] = toAdd2
    print("Title Sucsess")
except:
    print("Title Failed")

try:
    CheckedPage["Top 5 KW and Counts"] = toAdd3
    print("Top 5 KW and Counts Sucsess")
except:
    print("Top 5 KW and Counts Failed")

try:
    CheckedPage["server heads"] = toAdd4
    print("server heads Sucsess")
except:
    print()

try:
    CheckedPage["Contains bad words"] = toAdd5
    print("Contains bad words Sucsess")
except:
    print("Contains bad words Fail")

try:
    CheckedPage["Word Count"] = toAdd6
    print("Word Count Sucsess")
except:
    print("Word Count Fail")

try:
    CheckedPage["Sem ranks"] = toAdd7
    print("Sem ranks")
except:
    print("Sem ranks")

try:
    CheckedPage["Org KWs"] = toAdd8
    print("Org KWs Sucsess")
except:
    print("Org KWs Fail")

try:
    CheckedPage["OrgTrafs"] = toAdd9
    print("OrgTrafs Sucsess")
except:
    print("OrgTrafs Fail")

try:
    CheckedPage["Organic Costs"] = toAdd10
    print("Organic Costs Sucsess")
except:
    print("Organic Costs Fail")

try:
    CheckedPage["Containts Bad words"] = toAdd11
    print("Containts naughty words Sucsess")
except:
    print("Containts naughty words Fail")


CheckedPage.to_excel(ExOut)