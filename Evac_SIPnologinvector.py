"""
Script created to scrape pull the badge logs from a SaaS website The Company used.
 After running the roll call it will use Excel to generate an Accountability report during a Shelter in Place emergency. 
 Reports can be saved to the Teams site
 
 Note: This script has been modified from its fully working version to remove any confidential information.
    It does not have the ability to webdrive or save to a network folder.
 """

import xlwings as xw
import pandas as pd
import re
import numpy as np
#import win32com.client as win32
#from selenium import webdriver
#from selenium.webdriver.common.keys import Keys
#from selenium.webdriver.support.ui import Select
import time
from datetime import timedelta
import sys
import requests
import json
#from requests_negotiate_sspi import HttpNegotiateAuth
#from win32com.client import Dispatch

xw.Book('Evac_Report_Toolvec.xlsm').set_mock_caller()
wb = xw.Book.caller()
wbs = wb.sheets['StartHere']
wbf = wb.sheets['FacilitiesReportPaste']
wbsh = wb.sheets['SIPEvacReportPaste']
wbnm = wb.sheets['NeverMustered']
wbai = wb.sheets['Badged After Incident']
wbmu = wb.sheets['Mustered']
wbmue = wb.sheets['Mustered Too Early']
wbnbd = wb.sheets['No Badge Data']
wbsec = wb.sheets['SecurityEntry']
wbert = wb.sheets['ERT List']
wbr = wb.sheets['Roster']
sitename = wbs.range('B7').value #Site selected in Excel sheet


######################################
#Commenting out below code that was used to scrape the badge reader online log system. This would requires access to The Company's
    #system that is not possible or permitted.
#The accompanying spreadsheet has dummy data already added.

'''
#Declaring login variables for the log site. The Company has 2 different plants that this report can be run for.
un = 'fake' #username
up = 'password' #password

if sitename == "Plant1":
    plant = 'PlantCode=X002'
    facname = "19"
    sipname = "97"
else:
    plant = 'PlantCode=X006'
    facname = "18"
    sipname = "80"

#19 = Plant1 Facility, 97 = Plant1 Shelter in Place Report, 18 = Plant2 Facility, 80 = Plant2 SIP
print("Please wait for the program to open the browser, log in to LogSite, and run the roll call reports.")

options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging']) #Removing webdriver error log printing
parser = Dispatch("Scripting.FileSystemObject")
chrome_browser = 'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe' #--Chrome.exe filepath
version = parser.GetFileVersion(chrome_browser) #Checks chrome version
chrome_browser_version = version[:2] # substring version to capabable version (ie. 100)
driver_loc = "\\\\global.widget404.com\\network\\Shared\\Accountability\\WebDrivers" #network drive location webdrive is saved
currentPath = driver_loc + "\\chromedriver_"+chrome_browser_version + '.exe'
driver = webdriver.Chrome(currentPath, options=options)# Using Chrome to access web

#chromedriver exe file needs to be saved to a folder anybody running the script can access
#chromedriver can be downloaded from https://chromedriver.chromium.org/downloads
#if the script fails at the download step, make sure the version of chrome on your computer matches the version of chromedrive downloaded (ex v100) """

time.sleep(5)

# Open the LogSite website using chromedriver
driver.get('http://x0kuxps2tx01.widget404.com/login/') #Website that houses badge reader logs.
window_before = driver.window_handles[0]

# Locate id and password boxes on site
id_box = driver.find_element_by_css_selector('input#username')
pass_box = driver.find_element_by_css_selector('input#password')

# Fill in login information
id_box.send_keys(un) #input username
pass_box.send_keys(up) #input password
pass_box.send_keys(Keys.ENTER)

time.sleep(5)
driver.switch_to.frame("mainFrame")
iframe = driver.find_element_by_xpath("//iframe[@name='innerPageFrame']")
driver.switch_to.frame(iframe)

#Run Roll Call to see who is in the Facility - then it will place it in spreadsheet
site_select1 = Select(driver.find_element_by_id("regionid"))
site_select1.select_by_value(facname)

#selects child region and hits the Go button
driver.find_element_by_css_selector('input#recurse').click() #Select child region box
driver.find_element_by_css_selector('input#form0_save').click() #Click Go button
time.sleep(2)
driver.find_element_by_xpath("//*[@id='footer']/div/button").click() #Click refresh now button

time.sleep(2)

driver.switch_to.default_content()
driver.switch_to.frame("mainFrame")
driver.find_element_by_css_selector('div#topbarright').click() #Click menu button
driver.find_element_by_css_selector('div#print').click() #Click print report button

time.sleep(10)
#Copying the popup window then closing it
window_after = driver.window_handles[1]
driver.switch_to.window(window_after)

rollcall = driver.find_element_by_css_selector("body")
rollcall.send_keys(Keys.CONTROL,"a")
rollcall.send_keys(Keys.CONTROL,"c")
time.sleep(3)
driver.find_element_by_xpath("//*[@id='control']/button[2]").click() #close report print window
time.sleep(3)
wb.macro('Fpaste')() #Run macro to paste matching destination

#Running Roll Call on Shelter in Place
driver.switch_to.window(window_before)
driver.switch_to.frame("mainFrame")
driver.switch_to.frame(iframe)

site_select1 = Select(driver.find_element_by_id("regionid"))
site_select1.select_by_value(sipname)
driver.find_element_by_css_selector('input#form0_save').click()

time.sleep(2)
try:
    driver.find_element_by_xpath("//*[@id='footer']/div/button").click()
    time.sleep(2)
    driver.switch_to.default_content()
    driver.switch_to.frame("mainFrame")
    driver.find_element_by_css_selector('div#topbarright').click()
    driver.find_element_by_css_selector('div#print').click()
    time.sleep(10)
    #Copying the popup window then closing it
    window_after = driver.window_handles[1]
    driver.switch_to.window(window_after)
    rollcall = driver.find_element_by_css_selector("body")
    rollcall.send_keys(Keys.CONTROL,"a")
    rollcall.send_keys(Keys.CONTROL,"c")
    time.sleep(3)
    driver.find_element_by_xpath("//*[@id='control']/button[2]").click()
    time.sleep(3)
    wb.macro('Spaste')()     #Pasting in shelter in place tab
    sipdf = pd.DataFrame(wbsh.range('A3').expand().options(numbers=int).value).drop([2], axis=1) #Create dataframe shelter in place report
    sipdf.columns = ['ID', 'Last Name', 'Region', 'Phone Number', 'Date/Time of Last Access', 'Reader', 'Status']
except:
    sipdf = pd.DataFrame(columns = ['ID', 'Last Name', 'Region', 'Phone Number', 'Date/Time of Last Access', 'Reader', 'Status'])
    print("There were no sheltered persons.")
driver.switch_to.window(window_before)
driver.quit() #Closes webdrive
'''

##############################################################################



#Create dataframes of facility report ('fac') and shelter in place ('sip) report sheets
facdf = pd.DataFrame(wbf.range('A3').current_region.options(numbers = int).value).drop([2],axis =1).iloc[2:]
    #current_region selects all cells on the page.
    #drop[2] removes a "ID Photo" column that is not used.
    #.iloc[2:] removes the first two rows. These are pasted values that are not necessary.
    #THERE IS FOR SURE A BETTER WAY TO DO THIS!

headsdf = facdf.iloc[0] #Creating headers for all of the dataframes and tabs using the first row.

facdf = facdf[1:] #removes the redundant header row from the dataframe.
facdf.columns = headsdf #setting the column names based on the extracted headers.
facdf = facdf[facdf['Last Name'].str.len() > 3] #Removing any names that are less than 3 characters. Junk entries sometimes show up.
facdf = facdf[facdf["Region"].notnull()].sort_values(by='Last Name') #sorting by last name for processing speed and for readability.
facdf["Date/Time of Last Access"] = pd.to_datetime(facdf["Date/Time of Last Access"], format = '%-m/-%d/%y %-I:%M', errors = 'coerce') #Forcing pandas to read this as a datetime.

sipdf = pd.DataFrame(wbsh.range('A3').current_region.options(numbers = int).value).drop([2],axis =1).iloc[2:]
sipdf = sipdf[1:]
sipdf.columns = headsdf
sipdf = sipdf[sipdf['Last Name'].str.len() > 3]
sipdf = sipdf[sipdf["Region"].notnull()].sort_values(by='Last Name')
sipdf["Date/Time of Last Access"] = pd.to_datetime(sipdf["Date/Time of Last Access"], format = '%-m/-%d/%y %-I:%M', errors = 'coerce')

starttime = pd.to_datetime(wbs.range('A9').value) #Start time of the emergency.
oldtime = pd.to_datetime(wbs.range('B9').value) #The last time in the past before the emergency that badge access data should be considered valid.

# ###################################################################################
# This old code block uses functions and apply to create a "Status" column. I've since learned that vectorizing is faster. I'm keeping it just because.
# def FacStatusCheck(ID, access):
#     if pd.isna(ID) or pd.isna(access): #No badge ID or last time badge used data.
#         return "No Badge Data"
#     if access > starttime: #If they came on site after the emergency, but haven't mustered. They may not have heard the alarm.
#         return "Badged After Incident"
#     if access < oldtime: #If the last time they badged indicates the data is old and they are not actually on site.
#             return "Old Badge Data"
#     return "Never Mustered" #These are the people that Emergency Response should be looking for.
# def SIPStatusCheck(ID, access):
#     if pd.isna(ID) or pd.isna(access):
#         return "No Badge Data"
#     if access > starttime:
#         return "Mustered" #Mark as safe as they have checked in at a muster reader.
#     return "Mustered Too Early" #Scanned a muster reader before an emergency. Usually because they were anticipating a drill.
# facdf['Status'] = facdf.apply(lambda x: FacStatusCheck(x["ID"],x["Date/Time of Last Access"]),axis=1)
# sipdf['Status'] = sipdf.apply(lambda x: SIPStatusCheck(x["ID"],x["Date/Time of Last Access"]),axis=1)
# ###################################################################################

######### Vectorized version to create a "Status" column with the condition of each row. ############
conds = [ #conditions are evaluated in order
    facdf['ID'].isna() | facdf['Date/Time of Last Access'].isna(), #No badge ID or last time badge used data.
    facdf['Date/Time of Last Access'] > starttime, #If they came on site after the emergency, but haven't mustered. They may not have heard the alarm.
    facdf['Date/Time of Last Access'] < oldtime, #If the last time they badged indicates the data is old and they are not actually on site.
]
choices = [ #matches condition to what will be added to the column.
    'No Badge Data',
    'Badged After Incident',
    'Old Badge Data'
]
facdf['Status'] = np.select(conds, choices, default='Never Mustered') #These are the people that Emergency Response should be looking for.

conds = [
    sipdf['ID'].isna() | sipdf['Date/Time of Last Access'].isna(), #No badge ID or last time badge used data.
    sipdf['Date/Time of Last Access'] > starttime #Mark as safe as they have checked in at a muster reader.
]
choices = [
    'No Badge Data',
    'Mustered'
]

sipdf['Status'] = np.select(conds, choices, default='Mustered Too Early') #Scanned a muster reader before an emergency. Usually because they were anticipating a drill.

###################################################################################################

nmdf = facdf.loc[facdf.Status == "Never Mustered"].drop('Status', axis = 1) #never mustered dataframe
aidf = facdf.loc[facdf.Status == "Badged After Incident"].drop('Status', axis = 1) #badged after incident
nbdf = facdf.loc[facdf.Status == "No Badge Data"].drop('Status', axis = 1) #no badge data in unmustered report

ncount = len(nmdf) #counting number of people that did not muster.
aicount = len(aidf) #Counting the number of people that badged into the site after the emergency, and haven't mustered.

try: #Doing error handling if reporting tool is run without anybody sheltering
    mdf = sipdf.loc[sipdf.Status == "Mustered"].drop('Status', axis = 1) #mustered dataframe
    mtdf = sipdf.loc[sipdf.Status == "Mustered Too Early"].drop('Status', axis = 1) #mustered before incident time
    nbdf2 = sipdf.loc[sipdf.Status == "No Badge Data"].drop('Status', axis = 1) #no badge data in SIP report
    nbdf = pd.concat([nbdf,nbdf2]) #combining no badge data
    mcount = len(mdf) #Counting number of people that mustered.
    mtcount = len(mtdf) #Counting number of people that mustered too early.
except:
    mdf = None
    mtdf = None
    mcount = 0
    mtcount = 0

nbdcount = len(nbdf) #Counting number of people in category

try:
    #Pulling the visitor logs from Security. It's in a json format that will be converted to a dataframe.
    # print('Downloading Security Entry and Truck Yard logs and pasting them in.')
    # base_url = 'http://xz.widget404.com/plantview/StreamEntryLog?' #SaaS site base URL.
    # url_tail = "&OnsiteOnly=true&Columns=TimeIn,Gate,EntryType,EmployeeID,Name,Company,ContactNumber,VehicleNumber,LicenseNumber,Visiting,Comments,Badge" #adding columns to pull

    # response = requests.get(base_url + plant + url_tail, auth=HttpNegotiateAuth()) #combining to make full URL + using credentials to authorize.
    # response.raise_for_status()

    # secpull = json.loads(response.content)
    # secdf = pd.json_normalize(secpull)
    # secdf['TimeIn'] = pd.to_datetime(secdf['TimeIn'])
    # wbsec.range('A1').options(index=False, header=True).value = secdf 
    
    ########################################################################################
    #This section was added for demo purposes. Dummy entries are already in the sheet.
    secdf = pd.DataFrame(wbsec.range('A1').current_region.options(numbers = int).value)
    headsdf = secdf.iloc[0] #Creating headers for all of the dataframes and tabs using the first row.
    secdf = secdf[1:] #removes the redundant header row from the dataframe.
    secdf.columns = headsdf #setting the column names based on the extracted headers.
    ########################################################################################
    
    secdf['Badge'] = secdf['Badge'].astype(str).str.strip()
    secdf['ContactNumber'] = secdf['ContactNumber'].astype(str).str.strip()
    secdf['Name'] = secdf['Name'].astype(str).str.strip()
    secdf = secdf.drop_duplicates(subset='Badge', keep='first')

    #Attempts to match visitor badges phone numbers to never mustered, mustered early, or badged after incident.
    #Otherwise the SaaS dump would just show a visitor badge number.
    print('Attempting to match unaccounted visitor badges to phone numbers')

    badgeid_to_phone = secdf.set_index('Badge')['ContactNumber'].to_dict()
    badgeid_to_name  = secdf.set_index('Badge')['Name'].to_dict()
    
    pd.set_option('future.no_silent_downcasting', True) #Fixes a deprecation error I encountered.
    
    def replace_from_seclog(df):
        dfid = df['ID'].astype(str).str.strip()
        # map -> returns NaN when no mapping exists; fillna keeps original value
        df['Phone Number'] = dfid.map(badgeid_to_phone).fillna(df['Phone Number'])
        df['Last Name'] = dfid.map(badgeid_to_name).fillna(df['Last Name'])
        return df

    # apply to your frames
    nmdf = replace_from_seclog(nmdf)
    mtdf = replace_from_seclog(mtdf)
    aidf = replace_from_seclog(aidf)

except:
    print("Either there were no Security/Truck entries or there was an error. Please check PlantView.")
    pass

try:
#     #Opening up ERT Roster to find names of ERT members and their codes
    print("Checking ERT Roster to find names of members that are on site, but not sheltered")
#     wbo = xw.books.open("\\\\global.widget404.com\\network\\Plant1 Facilities Protection\\MiscDocs\\ERT Roster.xlsx")
#     time.sleep(2)
#     wbop = wbo.sheets['Personnel']

#     #Pasting current ERT Roster to report file
#     wbr.range('A1').options(index=False, header=False).value = wbop.range('B1:B300').options(numbers = int, ndim = 2).value #Name
#     wbr.range('B1').options(index=False, header=True).value = wbop.range('V1:V300').options(numbers = int, ndim = 2).value #Code
#     wbr.range('C1').options(transpose = True).value = "Present", " "
#     wbr.autofit()

######### Dummy data already in sheet ################

    #Creating dataframe for roster
    perdf = pd.DataFrame(wbr.range('A1').expand().value)
    perdf.columns = ['Name','Codes','Present']
    perdf['Name']=perdf['Name'].str.upper()
#     wbr.range('A1').options(index=False, header=False).value = perdf

    #Searching roll calls for ERT members on site
    for line, row in enumerate(facdf.itertuples(),1):
        regex = re.compile(row[2], re.IGNORECASE)
        matched = [x for x in perdf['Name'].values if regex.match(x)]
        if matched == []:
            continue
        else:
            facdf.at[row.Index, 'Present'] = "Yes"
            facdf.at[row.Index, 'Roster Name'] = perdf.loc[perdf.Name == matched[0], 'Name'].values[0]

    #Creating dataframe of on site ERT members
    ertdf = facdf.loc[facdf.Present == "Yes"]
    ertdf = ertdf.loc[:,['Roster Name']]
    ertdf = ertdf.drop_duplicates(subset = 'Roster Name')
    wbert.range('A1').options(index=False, header=True).value = ertdf.sort_values(by="Roster Name")
#     wbo.close() #Closes roster file
except:
    # wbo.close()
    print("Something went wrong with getting the ERT roster. Skipping")
    pass

#Pasting filtered lists to respective tabs
wbnm.range('A1').options(index=False, header=True).value = nmdf.sort_values(by = 'Last Name')
wbai.range('A1').options(index=False, header=True).value = aidf.sort_values(by = 'Last Name')
wbmu.range('A1').options(index=False, header=True).value = mdf.sort_values(by = 'Last Name')
wbmue.range('A1').options(index=False, header=True).value = mtdf.sort_values(by = 'Last Name')
wbnbd.range('A1').options(index=False, header=True).value = nbdf.sort_values(by = 'Last Name')

#Format column width
wbnm.autofit()
wbai.autofit()
wbmu.autofit()
wbmue.autofit()
wbnbd.autofit()
wbsec.autofit()
wbert.autofit()

wbs.range('B12').options(transpose=True).value = mcount, mtcount, ncount, aicount, nbdcount #Adding counts of each category to main page.
wbs.range('B20').value = time.strftime("%m/%d/%Y %H:%M") #Adding a timestamp of when the report was run.

# #Run save macro
# print('Answer save prompt in Excel. You may close this window after doing so.')
# wb.macro('SaveReport')()

# sys.exit()