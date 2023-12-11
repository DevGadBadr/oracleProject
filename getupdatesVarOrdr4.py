from selenium.webdriver.common.by import By
import time 
import bs4
from selenium.common.exceptions import NoSuchElementException,ElementClickInterceptedException,StaleElementReferenceException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from writetoexcelVarOrders import writetoexcelfn,writenextsteps
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService # Similar thing for firefox also!
from subprocess import CREATE_NO_WINDOW

def varordupdate(hidechrome,user,passw):
   
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    
    chrome_service = ChromeService('chromedriver.exe')
    chrome_service.creation_flags = CREATE_NO_WINDOW
    
    if hidechrome:
        driver = webdriver.Chrome(options=options,service=chrome_service)
    else:
        driver = webdriver.Chrome(service=chrome_service)

    getupdatesVarOrd(driver,user,passw)

def getupdatesVarOrd(driver,user,passw):
   
    print('Running Main Update Code for variation orders')
    start_time = time.time()
    writefn = writetoexcelfn
    
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
            
  
    #Choosing variation orders from left side
    counter = 0
    counter2=0
    time.sleep(8)
    loaded=False
    while not loaded:
        try:
            c=driver.find_element(By.XPATH,'//*[@id="subcontractor-contracts-menu"]/a[5]')
            c.click()
            loaded=True
            break
        except:
            time.sleep(1)
            counter +=1
            counter2+=1
            print(f'Choosing Variation Orders from left side {counter}')
            
        if counter==30:
            counter=0
            driver.refresh()
            time.sleep(6)
            #Switching to the frame
            while True:
                try:
                    frame = driver.find_element(By.ID,'pt1:_FOr1:0:_FONSr2:0:_FOTsr1:0:S_sdf_1596967376717_IF::f')
                    driver.switch_to.frame(frame)
                    break
                except NoSuchElementException:
                    time.sleep(1)
                    print('Switching to frame')

        if counter2==50:
            print('Internet connection problem')
            raise SystemExit('Disconnected')


    # Choosing project BU from dropdown menu
    counter=0
    while True:
        try:
            projectsbu = driver.find_element(By.XPATH,'//*[@id="inputState"]/option[2]')
            projectsbu.click()
            break
        except:
            counter+=1
            time.sleep(1)
            print('trying choose bu from menu')
        if counter==40:
            raise SystemExit('Aborted')
   

    #clicking yes to confirmation message for projects BU
    
    while True:
        try:
            time.sleep(2)    
            confirm = driver.find_element(By.XPATH,'//*[@id="cdk-overlay-0"]/nz-modal/div/div[2]/div/div/div/div/div[2]/button[2]')
            confirm.click()
            break
        except ElementClickInterceptedException:
            counter+=1
            time.sleep(1)
            print('Trying click yes')
        if counter==35:
            raise SystemExit('Aborted')
        
        
    #Reading Table Data
    trial=0
    counter=0
    while True:
        try:    
            driver.find_element(By.XPATH,'//*[@id="main-content"]/app-main-sb-vo/section/use-vo-list/app-subcontractor-variation-order-list/nz-table/nz-spin/div/div/div/div/table/tbody/tr[10]/td[4]')    
            break
        except ElementClickInterceptedException:
            time.sleep(1)
            print('Trying on table')
            trial+=1
            counter+=1
        except NoSuchElementException:
            time.sleep(1)
            print('Trying on table')
            trial+=1
            counter+=1
        if counter==40:
            raise SystemExit('Aborted')
        
        if trial==35:
            
            driver.refresh()
            time.sleep(5)
            trial=0
            # Switching to the new frame
            passed = False
            while not passed:
                try:
                    frame = driver.find_element(By.ID,'pt1:_FOr1:0:_FONSr2:0:_FOTsr1:0:S_sdf_1596967376717_IF::f')
                    driver.switch_to.frame(frame)
                    break
                except NoSuchElementException:
                    time.sleep(1)
                    print('Trying to switch to the frame level Choosing Variation Orders')
            #Choosing Variation orders from left side
            counter = 0
            counter2=0
            time.sleep(8)
            loaded=False
            while not loaded:
                try:
                    c=driver.find_element(By.XPATH,'//*[@id="subcontractor-contracts-menu"]/a[5]')
                    c.click()
                    loaded=True
                    break
                except:
                    time.sleep(1)
                    counter +=1
                    counter2+=1
                    print('Choosing Variation Orders from left side')
                    
                if counter==25:
                    counter=0
                    driver.refresh()
                    time.sleep(6)
                    #Switching to the frame
                    while True:
                        try:
                            frame = driver.find_element(By.ID,'pt1:_FOr1:0:_FONSr2:0:_FOTsr1:0:S_sdf_1596967376717_IF::f')
                            driver.switch_to.frame(frame)
                            break
                        except NoSuchElementException:
                            time.sleep(1)
                            print('Switching to frame')

                if counter2==40:
                    print('Internet connection problem')
                    raise SystemExit('Disconnected')
           
            
    
    # Getting the variation order table data
    three_pages_data=[]

    for n in range(2,5):
        
        #Click on page number
        clicked=False
        while not clicked:
            try:
                WebDriverWait(driver,10).until(EC.element_to_be_clickable((By.XPATH,f'//*[@id="main-content"]/app-main-sb-vo/section/use-vo-list/app-subcontractor-variation-order-list/nz-table/nz-spin/div/nz-pagination/ul/li[{n}]')))
                confirm = driver.find_element(By.XPATH,f'//*[@id="main-content"]/app-main-sb-vo/section/use-vo-list/app-subcontractor-variation-order-list/nz-table/nz-spin/div/nz-pagination/ul/li[{n}]')
                confirm.click()
                break
            except ElementClickInterceptedException:
                pass
        
        #waiting for data to load  
        loaded=False
        while not loaded:
            try:    
                driver.find_element(By.XPATH,'//*[@id="main-content"]/app-main-sb-vo/section/use-vo-list/app-subcontractor-variation-order-list/nz-table/nz-spin/div/div/div/div/table/tbody/tr[5]/td[4]')
                break
            except:
                time.sleep(1)
                pass
            
            
        #Making sure all data is loaded
        while True:
            try:
                driver.maximize_window()    
                time.sleep(5)
                driver.find_element(By.XPATH,'//*[@id="main-content"]/app-main-sb-vo/section/use-vo-list/app-subcontractor-variation-order-list/nz-table/nz-spin/div/div/div/div/table/tbody/tr[10]/td[4]')
                for n in range(1,11):
                    l=driver.find_element(By.XPATH,f'//*[@id="main-content"]/app-main-sb-vo/section/use-vo-list/app-subcontractor-variation-order-list/nz-table/nz-spin/div/div/div/div/table/tbody/tr[{n}]/td[4]')
                    if len(l.text)==0:
                        raise Exception('No Name')
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
                #Click on variation requests
                time.sleep(2)
                driver.execute_script("window.scrollTo(0, 0);")
                while True:
                    try:
                        c=driver.find_element(By.XPATH,'//*[@id="subcontractor-contracts-menu"]/a[2]')
                        c.click()
                        break
                    except ElementClickInterceptedException:
                        print('Clicking Variation Requests')
                
                time.sleep(3)
                #Click again on Variation order
                
                c=driver.find_element(By.XPATH,'//*[@id="subcontractor-contracts-menu"]/a[5]')
                c.click()
                
                time.sleep(3)
                
                print('Switched Back')
                ex=0
                #In case after reload there still no data show try until data is visible
                while True:
                    try:
                        for n in range(1,11):
                            time.sleep(5)
                            l=driver.find_element(By.XPATH,f'//*[@id="main-content"]/app-main-sb-vo/section/use-vo-list/app-subcontractor-variation-order-list/nz-table/nz-spin/div/div/div/div/table/tbody/tr[{n}]/td[4]')
                            if len(l.text)==0:
                                print(f'Element number {n} is Empty')
                                driver.refresh()
                                time.sleep(10)
                                #Switching to the frame
                                while True:
                                    try:
                                        frame = driver.find_element(By.ID,'pt1:_FOr1:0:_FONSr2:0:_FOTsr1:0:S_sdf_1596967376717_IF::f')
                                        driver.switch_to.frame(frame)
                                        break
                                    except NoSuchElementException:
                                        time.sleep(1)
                                        print('Trying to switch to the frame level 2')
                                print('Sleeping again')
                                #Clicking Variation Order
                                while True:
                                    try:
                                        c=driver.find_element(By.XPATH,'//*[@id="subcontractor-contracts-menu"]/a[5]')
                                        c.click()
                                        break
                                    except:
                                        time.sleep(2)
                                        print('Clicking Variation order')
                                time.sleep(5)              
                            else:
                                ex=1
                                break
                        if ex==1:
                            break
                    except NoSuchElementException:
                        pass
                n=n-1
                break
            
            
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
        
    #Writing data to Excel
    writefn(inprocess,rejected,approved,closed,draft)
    print('Data Saved\n')
    m=0
    nextsteps=[]
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
                        
                        
                #Choosing Varitaion orders from left side
                time.sleep(2)            
                loaded=False
                while not loaded:
                    try:
                        c=driver.find_element(By.XPATH,'//*[@id="subcontractor-contracts-menu"]/a[5]')
                        c.click()
                        loaded=True
                        break
                    except:
                        time.sleep(2)
                        print('Choosing Varitaion orders from left side')

                #Click the search box 
              

                while True:
                    try:
                        time.sleep(1)
                        c=driver.find_element(By.XPATH,'//*[@id="main-content"]/app-main-sb-vo/section/use-vo-list/app-subcontractor-variation-order-list/div/div[2]/div[1]/ng-select/div/div/div[2]')
                        c.click()
                        
                        html = driver.page_source
                        soup2 = bs4.BeautifulSoup(html,'lxml')
                        #Getting ID for inprocess option
                        drop = soup2.select('.ng-option')
                        id = drop[3].attrs['id']
                        break
                    except IndexError:
                        print('ID')
                        time.sleep(2)
                        pass

                time.sleep(1)
                c=driver.find_element(By.ID,id)
                c.click()

                c=driver.find_element(By.XPATH,'//*[@id="main-content"]/app-main-sb-vo/section/use-vo-list/app-subcontractor-variation-order-list/div/div[2]/div[4]/button')
                c.click()

                html = driver.page_source
                soup = bs4.BeautifulSoup(html,'lxml')

                ctr_inprocess = soup.select('.ant-table-row')[1:]
                numofctrinprocess= len(ctr_inprocess)
                break
            except NoSuchElementException:
                print('Trying again to make the search')
                
        # Enter the Variation Order
        while True:
            try:   
                c=driver.find_element(By.XPATH,f'//*[@id="main-content"]/app-main-sb-vo/section/use-vo-list/app-subcontractor-variation-order-list/nz-table/nz-spin/div/div/div/div/table/tbody/tr[{m}]/td[2]')
                c.click()
                break
            except StaleElementReferenceException:
                print('Entering Variation Order')
                
                
            except ElementClickInterceptedException:
                print('Entering Variation Order')
            
        # Click on History
        while True:  
            try:   
                c=driver.find_element(By.XPATH,f'//*[@id="main-content"]/use-vo-details/ul/li[5]/a')
                c.click()
                break
            except StaleElementReferenceException:
                print('Clikcing History')
                
        time.sleep(2)
        #Getting Next step from history
        while True:
            try:
                # Getting page source
                time.sleep(2)
                html = driver.page_source
                soup = bs4.BeautifulSoup(html,'lxml')
                if len(soup.select('use-history')[0].select('.text-primary'))==0:
                    print('This project is completed')
                    break
                nextstep = soup.select('use-history')[0].select('.text-primary')[-1].parent.text
                nextsteps.append(nextstep)
                print(f'Variation Order {n[0]} status:\n\n {nextstep}\n')
                break
            except IndexError:
                time.sleep(2)
                print('Getting Next step from History')
        
    
    writenextsteps(inprocess,nextsteps)      
    end_time = time.time()
    elapsed_time = end_time-start_time
    if elapsed_time>60:
        elapsed_time=elapsed_time/60
        print(f'Time taken to refresh data: {round(elapsed_time,2)} minutes')  
    else:
        print(f'Time taken to refresh data: {round(elapsed_time,2)} Seconds')  
