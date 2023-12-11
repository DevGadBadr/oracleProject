from selenium import webdriver
from selenium.webdriver.common.by import By
import time 
import bs4
import datetime
from selenium.common.exceptions import NoSuchElementException,ElementClickInterceptedException,StaleElementReferenceException,TimeoutException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from writetoexcelVarReq import writetoexcelfn,writenextsteps
from selenium.webdriver.support.select import Select
from PyQt5.QtCore import pyqtSignal,QThread,QObject
from selenium.webdriver.chrome.service import Service as ChromeService # Similar thing for firefox also!
from subprocess import CREATE_NO_WINDOW

def varrequpdate(hidechrome,user,passw):
    x= variationrequests()
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    
    chrome_service = ChromeService('chromedriver.exe')
    chrome_service.creation_flags = CREATE_NO_WINDOW
    
    if hidechrome:
        driver = webdriver.Chrome(options=options,service=chrome_service)
    else:
        driver = webdriver.Chrome(service=chrome_service)

    x.getupdatesvr(driver,user,passw)
    
    
class variationrequests(QObject):
    
    cprogress=0
    progress = pyqtSignal(int)
    
    def getupdatesvr(self,driver,user,passw):
      
        #Just a progress Update 5
        (5)
        #Just a progress Update   
                
        print('Running Main Update Code for Variation Requests')
        start_time = time.time()
        date = datetime.date.today()
        writefn = writetoexcelfn
        writestatus=writenextsteps

        url = 'https://fa-elto-saasfaprod1.fa.ocs.oraclecloud.com/fscmUI/faces/FuseWelcome?_afrLoop=67255470635805936&_afrWindowMode=0&_afrWindowId=null&_adf.ctrl-state=xvbh37olx_1&_afrFS=16&_afrMT=screen&_afrMFW=1036&_afrMFH=659&_afrMFDW=1536&_afrMFDH=864&_afrMFC=8&_afrMFCI=0&_afrMFM=0&_afrMFR=120&_afrMFG=0&_afrMFS=0&_afrMFO=0'

        options=webdriver.ChromeOptions()
        options.add_argument('--headless')

        #Just a progress Update 10
        (10)
        #Just a progress Update
            
        driver.get(url)

        # Sign in to Oracle
        id = driver.find_element(By.XPATH,'//*[@id="userid"]')
        id.send_keys(user)

        pas= driver.find_element(By.XPATH,'//*[@id="password"]')
        pas.send_keys(passw)

        signin = driver.find_element(By.XPATH,'//*[@id="btnActive"]')
        signin.click()

        #Just a progress Update 15
        (25)
        #Just a progress Update

        #Choose contracts in landing page
        passed = False
        while not passed:
            try:
                contracts = driver.find_element(By.XPATH,'//*[@id="groupNode_NewPages"]')
                contracts.click()
                contr = driver.find_element(By.XPATH,'//*[@id="EXT_EXTN1596967375013_MENU_1596967375652"]')
                contr.click()
                break
            except NoSuchElementException:
                time.sleep(1)
                print('Trying 1')

        # Switching to the frame
        passed = False
        while not passed:
            try:
                frame = driver.find_element(By.ID,'pt1:_FOr1:1:_FONSr2:0:_FOTsr1:0:S_sdf_1596967376717_IF::f')
                driver.switch_to.frame(frame)
                break
            except NoSuchElementException:
                time.sleep(1)
                print('Switching to frame')
                
        #Just a progress Update 20
        (30)
        #Just a progress Update


        #Choosing Variation requests from left side
        time.sleep(8)
        slowcount=0
        loaded=False
        while not loaded:
            if slowcount==40:
                driver.refresh()
                slowcount=0
                time.sleep(8)
            try:
                c=driver.find_element(By.XPATH,'//*[@id="subcontractor-contracts-menu"]/a[4]')
                c.click()
                loaded=True
                break
            except:
                time.sleep(1)
                slowcount+=1
                print('Choosing Variation requests from left side')

        #Just a progress Update 25
        (45)
        #Just a progress Update

        # Choosing project BU from dropdown menu
        chosen_pros=False
        while not chosen_pros:
            try:
            
                projectsbu = driver.find_element(By.XPATH,'//*[@id="inputState"]/option[2]')
                projectsbu.click()
                break
            except:
                time.sleep(1)
                print('trying choose bu from menu')
                
        #Just a progress Update 30
        (50)
        #Just a progress Update

        #clicking yes to confirmation message for projects BU
        while True:
            try:
                time.sleep(2)    
                confirm = driver.find_element(By.XPATH,'//*[@id="cdk-overlay-0"]/nz-modal/div/div[2]/div/div/div/div/div[2]/button[2]')
                confirm.click()
                break
            except ElementClickInterceptedException:
                time.sleep(1)
                print('Trying click yes')
            
        #Waiting Until table is loaded
        counter = 0
        counter2=0
        while True:
            try:    
                driver.find_element(By.XPATH,'//*[@id="main-content"]/app-main-sb-vr/section/use-vr-list/app-subcontractor-variation-request-list/nz-table/nz-spin/div/div/div/div/table/tbody')    
                break
            except ElementClickInterceptedException:
                time.sleep(1)
                counter+=1
                counter2+=1
                print('Trying on table ElementClickInterceptedException')
            except NoSuchElementException:
                time.sleep(1)
                counter+=1
                counter2+=1
                print('Trying on table NoSuchElementException')
            if counter==25:
                driver.refresh()
                time.sleep(6)
                print('Connection is slow')
                # Switching to the frame
                while True:
                    try:
                        frame = driver.find_element(By.ID,'pt1:_FOr1:0:_FONSr2:0:_FOTsr1:0:S_sdf_1596967376717_IF::f')
                        driver.switch_to.frame(frame)
                        break
                    except NoSuchElementException:
                        time.sleep(1)
                        print('Switching to frame')
                #Choosing Variation requests from left side
                time.sleep(3)
                loaded=False
                while not loaded:
                    try:
                        c=driver.find_element(By.XPATH,'//*[@id="subcontractor-contracts-menu"]/a[4]')
                        c.click()
                        loaded=True
                        break
                    except:
                        time.sleep(1)
                        slowcount+=1
                        print('Choosing Variation requests from left side')
                
            if counter2==50:
                print('Aborted due to bad connection')
                raise SystemExit('Aborted')
                
        #Just a progress Update 35
        (60)
        #Just a progress Update

        # Getting the contracts table data
        three_pages_data=[]

        for n in range(2,5):
            
            #Click on page number
            clicked=False
            while not clicked:
                try:    
                    WebDriverWait(driver,15).until(EC.element_to_be_clickable((By.XPATH,f'//*[@id="main-content"]/app-main-sb-vr/section/use-vr-list/app-subcontractor-variation-request-list/nz-table/nz-spin/div/nz-pagination/ul/li[{n}]')))
                    confirm = driver.find_element(By.XPATH,f'//*[@id="main-content"]/app-main-sb-vr/section/use-vr-list/app-subcontractor-variation-request-list/nz-table/nz-spin/div/nz-pagination/ul/li[{n}]')
                    confirm.click()
                    break
                except ElementClickInterceptedException:
                    pass
                except TimeoutException:
                    print('Your Internet is so slow I will Try again')
                    pass
            
            #waiting for data to load  
            while True:
                try:    
                    driver.find_element(By.XPATH,'//*[@id="main-content"]/app-main-sb-vr/section/use-vr-list/app-subcontractor-variation-request-list/nz-table/nz-spin/div/div/div/div/table/tbody/tr[10]/td[4]')
                    break
                except:
                    time.sleep(1)
                    pass
                
            #Just a progress Update 40
            (70)
            #Just a progress Update
                
            #Making sure all data is loaded
            if slowcount>40:
                co=8
                print('You Connection is poor')
            else:
                co=4
            while True:
                try:    
                    time.sleep(co)
                    driver.find_element(By.XPATH,'//*[@id="main-content"]/app-main-sb-vr/section/use-vr-list/app-subcontractor-variation-request-list/nz-table/nz-spin/div/div/div/div/table/tbody/tr[1]/td[4]')
                    for n in range(1,11):
                        l=driver.find_element(By.XPATH,f'//*[@id="main-content"]/app-main-sb-vr/section/use-vr-list/app-subcontractor-variation-request-list/nz-table/nz-spin/div/div/div/div/table/tbody/tr[{n}]/td[4]')
                        if len(l.text)==0:
                            raise Exception('No Name for contractor')
                    break
                except ElementClickInterceptedException:
                    time.sleep(1)
                    print('Trying ElementClickInterceptedException')
                except NoSuchElementException:
                    time.sleep(1)
                    print('Trying NoSuchElementException')  
                    
                #Incase found a box with no data
                except Exception:
                    print('Exception has occured for empty data box')
                    driver.refresh()
                    time.sleep(5)
                    
                    # Switching to the new frame
                    passed = False
                    while not passed:
                        try:
                            frame = driver.find_element(By.ID,'pt1:_FOr1:0:_FONSr2:0:_FOTsr1:0:S_sdf_1596967376717_IF::f')
                            driver.switch_to.frame(frame)
                            break
                        except NoSuchElementException:
                            time.sleep(1)
                            print('Trying to switch to the frame level 1')
                    
                    while True:
                        try:
                            #Choosing Varition requests
                            print('Sleeping')
                            c=driver.find_element(By.XPATH,'//*[@id="subcontractor-contracts-menu"]/a[4]')
                            c.click()
                            time.sleep(3)
                            ex=0
                            break
                        except NoSuchElementException:
                            time.sleep(3)
                            print('Your Connection is very slow')
                    
                    #Just a progress Update 45
                    (75)
                    #Just a progress Update
                    
                    #In case after reload there still no data show try until data is visible
                    while True:
                        try:
                            for n in range(1,11):
                                l=driver.find_element(By.XPATH,f'//*[@id="main-content"]/app-main-sb-vr/section/use-vr-list/app-subcontractor-variation-request-list/nz-table/nz-spin/div/div/div/div/table/tbody/tr[{n}]/td[4]')
                                if len(l.text)==0:
                                    print(f'Element number {n} is empty')
                                    driver.refresh()
                                    time.sleep(5)
                                    # Switching to the frame
                                    while True:
                                        try:
                                            frame = driver.find_element(By.ID,'pt1:_FOr1:0:_FONSr2:0:_FOTsr1:0:S_sdf_1596967376717_IF::f')
                                            driver.switch_to.frame(frame)
                                            break
                                        except NoSuchElementException:
                                            time.sleep(1)
                                            print('Trying to switch to the frame level 2')
                                    print('Sleeping again')
                                    c=driver.find_element(By.XPATH,'//*[@id="subcontractor-contracts-menu"]/a[4]')
                                    c.click()
                                    time.sleep(3)
                                else:
                                    ex=1
                                    break
                            if ex==1:
                                break
                        except NoSuchElementException:
                            pass
                    n=n-1
                    break
                
            #Just a progress Update 50
            (80)
            #Just a progress Update
                
            #Getting data source code  
            html = driver.page_source
            soup = bs4.BeautifulSoup(html,'lxml')
            contracts = soup.select('.ant-table-row')[1:]
            
            #Filtering and sorting data         
            neat_contracts = []
            for contract in contracts:
                
                cont_data = contract.select('td')
                cont_data.pop(8)
                neat_contracts.append(cont_data)

            contracts_data=[]
            for cont in neat_contracts:
                contract=[]
                for data in cont:
                    t = data.text
                    contract.append(t)
                contracts_data.append(contract)   
            three_pages_data.append(contracts_data)
            
        #Just a progress Update 55
        (85)
        #Just a progress Update

        closed=[]
        approved=[]
        inprocess=[]
        rejected=[]  
        draft=[]     
        for page in three_pages_data:
            for contract in page:
                if contract[6] == 'Closed':
                    if contract in closed:
                        pass
                    else:
                        closed.append(contract)
                elif contract[6] == 'Approved':
                    if contract in approved:
                        pass
                    else:
                        approved.append(contract)
                elif contract[6] == 'InProcess':
                    if contract in inprocess:
                        pass
                    else:
                        inprocess.append(contract)
                elif contract[6] == 'Rejected':
                    if contract in rejected:
                        pass
                    else:
                        rejected.append(contract)
                elif contract[6] == 'Draft':
                    if contract in draft:
                        pass
                    else:
                        draft.append(contract)
                else:
                    print('There is something wrong') 
                    
        #Just a progress Update 60
        (90)
        #Just a progress Update
            
        #Writing data to Excel
        writefn(inprocess,rejected,approved,closed,draft)
        print('Data Saved\n')
        m=0
        nextsteps=[]
        #Getting Data for the inprocess tasks
        for n in inprocess:
            m+=1
            searched=False
            while not searched:
                try:
                    driver.refresh()
                    time.sleep(2)
                    # Switching to the frame
                    passed = False
                    while not passed:
                        try:
                            frame = driver.find_element(By.ID,'pt1:_FOr1:0:_FONSr2:0:_FOTsr1:0:S_sdf_1596967376717_IF::f')
                            driver.switch_to.frame(frame)
                            break
                        except NoSuchElementException:
                            time.sleep(1)
                            print('Trying to switch to the frame level after refresh')
                            
                    #Just a progress Update 65
                    (65)
                    #Just a progress Update
                            
                    #Choosing Payment Certificates from left side
                    time.sleep(2)            
                    loaded=False
                    while not loaded:
                        try:
                            c=driver.find_element(By.XPATH,'//*[@id="subcontractor-contracts-menu"]/a[4]')
                            c.click()
                            loaded=True
                            break
                        except:
                            time.sleep(2)
                            print('Choosing Payments Certificates from left side')
                            

                    #Click on drop down menu to choose the InProcess
                    while True:
                        try:
                            time.sleep(1)
                            c=driver.find_element(By.XPATH,'//*[@id="main-content"]/app-main-sb-vr/section/use-vr-list/app-subcontractor-variation-request-list/div/div[2]/div[1]/ng-select/div')
                            c.click()
                            break
                        except ElementClickInterceptedException:
                            print('Click intercepted Trying again')
                            driver.refresh()
                            time.sleep(5)
                                # Switching to the frame
                            passed = False
                            while not passed:
                                try:
                                    frame = driver.find_element(By.ID,'pt1:_FOr1:0:_FONSr2:0:_FOTsr1:0:S_sdf_1596967376717_IF::f')
                                    driver.switch_to.frame(frame)
                                    break
                                except NoSuchElementException:
                                    time.sleep(1)
                                    print('Trying to switch to the frame level after refresh')
                            #Choosing Payment Certificates from left side            
                            loaded=False
                            while not loaded:
                                try:
                                    c=driver.find_element(By.XPATH,'//*[@id="subcontractor-contracts-menu"]/a[4]')
                                    c.click()
                                    loaded=True
                                    break
                                except:
                                    time.sleep(2)
                                    print('Choosing Payments Certificates from left side')
                        
                    time.sleep(1)
                    html = driver.page_source
                    soup2 = bs4.BeautifulSoup(html,'lxml')
                    
                    while True:
                        try:
                            #Getting ID for inprocess option
                            drop = soup2.select('.ng-option')
                            id = drop[3].attrs['id']
                            break
                        except IndexError:
                            pass
                        
                    
                    #Just a progress Update 70
                    (92)
                    #Just a progress Update
                    
                    #choosing InProcess option
                    time.sleep(1)
                    c=driver.find_element(By.ID,id)
                    c.click()
                    
                    #Clicking on Search button
                    while True:
                        try:
                            c=driver.find_element(By.XPATH,'//*[@id="main-content"]/app-main-sb-vr/section/use-vr-list/app-subcontractor-variation-request-list/div/div[2]/div[4]/button')
                            c.click()
                            break
                        except ElementClickInterceptedException:
                            pass

                    html = driver.page_source
                    soup = bs4.BeautifulSoup(html,'lxml')

                    ctr_inprocess = soup.select('.ant-table-row')[1:]
                    numofctrinprocess= len(ctr_inprocess)
                    break
                except NoSuchElementException:
                    print('Trying again to make the search')
                    
            #Just a progress Update 75
            (95)
            #Just a progress Update

            # Enter the contract
            while True:
                try:   
                    c=driver.find_element(By.XPATH,f'//*[@id="main-content"]/app-main-sb-vr/section/use-vr-list/app-subcontractor-variation-request-list/nz-table/nz-spin/div/div/div/div/table/tbody/tr[{m}]/td[3]')
                    c.click()
                    break
                except StaleElementReferenceException:
                    print('Entering Payment Certificate')
                    
                except ElementClickInterceptedException:
                    print('Entering Payment Certificate')
                
            # Click on History
            while True:  
                try:   
                    time.sleep(2)
                    c=driver.find_element(By.XPATH,'//*[@id="main-content"]/use-vr-details/ul/li[5]/a')
                    c.click()
                    break
                except StaleElementReferenceException:
                    print('Clicking History')
                except NoSuchElementException:
                    print('No Element Exception Trying again')
                    
            #Just a progress Update 
            (96)
            #Just a progress Update
                    
            time.sleep(2)
            #Getting Next step from history
            while True:
                time.sleep(2)
                try:
                    c=driver.find_element(By.XPATH,'//*[@id="main-content"]/use-vr-details/section[2]/div/use-vr-history/use-history/app-no-data/div/h4')
                    if 'History' in c.text:   
                        continue
                except:
                    pass
                
                try:
                    # Getting page source
                    html = driver.page_source
                    soup = bs4.BeautifulSoup(html,'lxml')
                    if len(soup.select('use-history')[0].select('.text-primary'))==0:
                        print('This project is completed')
                        break
                    nextstep = soup.select('use-history')[0].select('.text-primary')[-1].parent.text
                    print(f'Payment Certificate {n[0]} status:\n\n {nextstep}\n')
                    nextsteps.append(nextstep)
                    break
                except IndexError:
                    time.sleep(2)
                    print('Getting Next step from History')
            
            #Just a progress Update 90
            (97)
            #Just a progress Update

        writestatus(nextsteps)

        end_time = time.time()
        elapsed_time = end_time-start_time
        if elapsed_time>60:
            elapsed_time=elapsed_time/60
            print(f'Time taken to refresh data: {round(elapsed_time,2)} minutes')  
        else:
            print(f'Time taken to refresh data: {round(elapsed_time,2)} Seconds')  

        #Just a progress Update 100
        (100)
        #Just a progress Update