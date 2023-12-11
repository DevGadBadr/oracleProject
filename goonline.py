from selenium.webdriver.common.by import By
import time 
import bs4
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException,ElementClickInterceptedException,StaleElementReferenceException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from writetoexcel import writetoexcelfn,writenextsteps
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService # Similar thing for firefox also!
from subprocess import CREATE_NO_WINDOW
from PyQt5.QtCore import pyqtSignal,QThread,QObject




def onlinefunc(hidechrome,user,passw):

    options = webdriver.ChromeOptions()
    options.add_argument('--headless')

    chrome_service = ChromeService('chromedriver.exe')
    chrome_service.creation_flags = CREATE_NO_WINDOW

    global driver
    if hidechrome:
        driver = webdriver.Chrome(options=options,service=chrome_service)
    else:
        driver = webdriver.Chrome(service=chrome_service)

    url = 'https://fa-elto-saasfaprod1.fa.ocs.oraclecloud.com/fscmUI/faces/FuseWelcome?_afrLoop=67255470635805936&_afrWindowMode=0&_afrWindowId=null&_adf.ctrl-state=xvbh37olx_1&_afrFS=16&_afrMT=screen&_afrMFW=1036&_afrMFH=659&_afrMFDW=1536&_afrMFDH=864&_afrMFC=8&_afrMFCI=0&_afrMFM=0&_afrMFR=120&_afrMFG=0&_afrMFS=0&_afrMFO=0'

    driver.get(url)

    # Sign in to Oracle
    id = driver.find_element(By.XPATH,'//*[@id="userid"]')
    id.send_keys(user)

    pas= driver.find_element(By.XPATH,'//*[@id="password"]')
    pas.send_keys(passw)

    signin = driver.find_element(By.XPATH,'//*[@id="btnActive"]')
    signin.click()
    
    #Choose contracts in landing page
    counter=0
    while True:
        try:
            contracts = driver.find_element(By.XPATH,'//*[@id="groupNode_NewPages"]')
            contracts.click()
            contr = driver.find_element(By.XPATH,'//*[@id="EXT_EXTN1596967375013_MENU_1596967375652"]')
            contr.click()
            break
        except NoSuchElementException:
            counter+=1
            time.sleep(1)
            print('Trying 1')
            
        if counter==35:
            raise SystemExit('Aborted')
            

    # Switching to the frame
    counter=0
    while True:
        try:
            frame = driver.find_element(By.ID,'pt1:_FOr1:1:_FONSr2:0:_FOTsr1:0:S_sdf_1596967376717_IF::f')
            driver.switch_to.frame(frame)
            break
        except NoSuchElementException:
            counter+=1
            time.sleep(1)
            print(f'Switching to frame {counter}')
            
        if counter==35:
            raise SystemExit('Aborted')
            
    print('Driver Ready')
    
    return driver