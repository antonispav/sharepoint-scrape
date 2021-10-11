from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

import os
import time
import shutil
import mimetypes
import getpass


def listFoldersFiles(prevFlag):                         
    folders = 0                                 
    files = 0                                   
                                       
    try:                                        
        #list all the folders                  
        header = driver.find_element_by_xpath('/html/body/form/div[12]/div/div[2]/div[2]/div[3]/div[1]/div/div/div/table/tbody/tr/td/table[2]/tbody')
    except Exception as ex: #the folder is empty
        subFolders.pop(-1)
        driver.find_element_by_css_selector('#DeltaPlaceHolderPageTitleInTitleArea > span:nth-child(1) > span:nth-child(1) > a:nth-child(1)').click()#return to previews page
        listFoldersFiles(prevFlag)

    allFiles = header.find_elements_by_xpath(".//tr")

    for tr in allFiles:
        
        a = tr.find_element_by_xpath('.//td[3]/div/a')

        apress = a.get_attribute('onmousedown')
        if 'VerifyFolderHref' in apress: #it's a folder
            if os.path.join(mainFolder, '/'.join(filter(None, subFolders)),a.text) in visited: # this folder is already visited
                continue
            #create a folder
            subFolders.append(a.text)
            if not os.path.exists(os.path.join(mainFolder, '/'.join(subFolders))):
                os.makedirs(os.path.join(mainFolder, '/'.join(subFolders)))

            folders += 1
            a.click()#enter the folder
            visited.append(os.path.join(mainFolder, '/'.join(subFolders))) #add the folder to the visited list
            listFoldersFiles(prevFlag) # read the content of the new folder
        else: #it's a file
            ahref = a.get_attribute('href')
            downloadFile(tr)
            files += 1
    if files > 0:
        moveFiles(os.path.join(mainFolder, '/'.join(subFolders))) 

    if driver.find_elements_by_css_selector('.ms-promlink-button'): #there is a next page
        if prevFlag:#finished. move to previews folder/ back button isn't working / finished with the second page
            prevFlag = False #a flag to know if we are in the first or the second page
            subFolders.pop(-1)
            driver.find_element_by_css_selector('#DeltaPlaceHolderPageTitleInTitleArea > span:nth-child(1) > span:nth-child(1) > a:nth-child(1)').click()#return to previews page
            listFoldersFiles(prevFlag)
        driver.find_element_by_css_selector('.ms-promlink-button').click()#go to next page
        prevFlag = True #a flag to know if we are in the first or the second page / finished with the first page
        listFoldersFiles(prevFlag)

    if folders == 0: # we visited all the subfolders
        try:
            subFolders.pop(-1)
        except Exception as ex1:#finished scraping -> quit
            print("Finished!")
            driver.quit()
            exit()
        driver.find_element_by_css_selector('#DeltaPlaceHolderPageTitleInTitleArea > span:nth-child(1) > span:nth-child(1) > a:nth-child(1)').click()#return to previews page
        listFoldersFiles(prevFlag)


def downloadFile(tr):
    
    more = WebDriverWait(tr, 20).until(EC.element_to_be_clickable((By.XPATH, './/td[4]')))
    more.click()

    moremore = WebDriverWait(more, 20).until(EC.element_to_be_clickable((By.XPATH, ".//a[@class='js-callout-action ms-calloutLinkEnabled ms-calloutLink js-ellipsis25-a']"))).click()

    listmenu = WebDriverWait(more, 20).until(EC.element_to_be_clickable((By.XPATH, ".//div/div/div/div[3]/span[2]/span/div/ul/li[@text='Download a Copy']"))).click()

def moveFiles(dest):
    while bool([name for name in os.listdir('downloads') if(name.split('.')[-1]=='part')]): #check every 2 seconds until all files are completely downloaded
        time.sleep(2)
    for downloadedFile in os.listdir('downloads'):
        #move the downloaded file to the appropriate folder
        downloadedFile = '/'+downloadedFile
        shutil.move('downloads'+downloadedFile, dest+downloadedFile)


mainFolder = 'SharePoint Documents'
if not os.path.exists(mainFolder):
                os.makedirs(mainFolder)
if not os.path.exists('downloads'):
                os.makedirs('downloads')
url = ''
prevFlag = False
subFolders = []
visited = []
downloadsPath = os.getcwd()+'/downloads'

#ask user for login credentials
user = input('Username: ')
passwd = getpass.getpass()

# Firefox and as a result selenium can't save files in multiple directories, only in a specified forder
profile = webdriver.FirefoxProfile()
profile.set_preference("browser.download.folderList", 2) #don't use default Downloads directory
profile.set_preference("browser.download.manager.showWhenStarting", False) #turn of showing download progress
profile.set_preference("browser.download.dir", downloadsPath) #set the directory for downloads
profile.set_preference("browser.helperApps.neverAsk.saveToDisk", ','.join(list(it for it in mimetypes.types_map.values()))) #automatically download files( all mime-type files )
profile.set_preference("plugin.disable_full_page_plugin_for_types", ','.join(list(it for it in mimetypes.types_map.values())))
profile.set_preference("pdfjs.disabled", True)
driver = webdriver.Firefox(firefox_profile=profile, executable_path=os.getcwd()+'/geckodriver')
driver.get(url)

# find the login fields
username = driver.find_element_by_xpath('//*[@id="ctl00_ctl00_NICEMasterPageBodyContent_SiteContentPlaceholder_txtFormsLogin"]')
password = driver.find_element_by_xpath('//*[@id="ctl00_ctl00_NICEMasterPageBodyContent_SiteContentPlaceholder_txtFormsPassword"]')
submit = driver.find_element_by_xpath('//*[@id="ctl00_ctl00_NICEMasterPageBodyContent_SiteContentPlaceholder_btnFormsLogin"]')

#type login credentials
username.send_keys(user)
password.send_keys(passwd)
submit.click()

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="WPQ2_ListTitleViewSelectorMenu_Container_surfaceopt0"]')))

listFoldersFiles(prevFlag)
