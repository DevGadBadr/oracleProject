#Preparing Excel File indices            
import openpyxl
import string

def writetoexcelfn(inprocess,rejected,approved,closed):
    
    a = list(string.ascii_uppercase)
    a1 = []
    for n in range(26):
        a1.append('A' + a[n])
        
    excelfile =  openpyxl.load_workbook('.\\Excel Sheets\\Contracts.xlsx')
        
    sheet= excelfile['Sheet1']
    
    #Deleting Previous Data
    cols = a+a1 
    for col in cols:
        for num in range(3,201):
            cell = col + str(num)
            sheet[cell] = ''
            
    #Deleting Previous Data InProcess
    cols = a+a1 
    sheet1=excelfile['InProcess']
    for col in cols:
        for num in range(3,201):
            cell = col + str(num)
            sheet1[cell] = ''
            
    #InProcess
    d=2
    
    for contract in inprocess:
        d=d+1
        #write contract
        cell = a[0] + str(d)
        sheet[cell] = contract[0]
        #write project
        cell = a[1] + str(d)
        sheet[cell] = contract[1]
        #subcotractor
        cell = a[2] + str(d)
        sheet[cell] = contract[2]
        #subcontractor site
        cell = a[3] + str(d)
        sheet[cell] = contract[3]
        #contract total
        cell = a[4] + str(d)
        sheet[cell] = contract[4]
        #date 
        cell = a[5] + str(d)
        sheet[cell] = contract[5]
        #status
        cell = a[6] + str(d)
        sheet[cell] = contract[6]

    #Approved
    d=2
    for contract in approved:
        d=d+1
        #write contract
        cell = a[0+7] + str(d)
        sheet[cell] = contract[0]
        #write project
        cell = a[1+7] + str(d)
        sheet[cell] = contract[1]
        #subcotractor
        cell = a[2+7] + str(d)
        sheet[cell] = contract[2]
        #subcontractor site
        cell = a[3+7] + str(d)
        sheet[cell] = contract[3]
        #contract total
        cell = a[4+7] + str(d)
        sheet[cell] = contract[4]
        #date 
        cell = a[5+7] + str(d)
        sheet[cell] = contract[5]
        #status
        cell = a[6+7] + str(d)
        sheet[cell] = contract[6]

    #Rejected
    d=2
    for contract in rejected:
        d=d+1
        #write contract
        cell = a[0+14] + str(d)
        sheet[cell] = contract[0]
        #write project
        cell = a[1+14] + str(d)
        sheet[cell] = contract[1]
        #subcotractor
        cell = a[2+14] + str(d)
        sheet[cell] = contract[2]
        #subcontractor site
        cell = a[3+14] + str(d)
        sheet[cell] = contract[3]
        #contract total
        cell = a[4+14] + str(d)
        sheet[cell] = contract[4]
        #date 
        cell = a[5+14] + str(d)
        sheet[cell] = contract[5]
        #status
        cell = a[6+14] + str(d)
        sheet[cell] = contract[6]

    #Closed
    d=2
    for contract in closed:
        d=d+1
        #write contract
        cell = a[0+21] + str(d)
        sheet[cell] = contract[0]
        #write project
        cell = a[1+21] + str(d)
        sheet[cell] = contract[1]
        #subcotractor
        cell = a[2+21] + str(d)
        sheet[cell] = contract[2]
        #subcontractor site
        cell = a[3+21] + str(d)
        sheet[cell] = contract[3]
        #contract total
        cell = a[4+21] + str(d)
        sheet[cell] = contract[4]
        #date 
        cell = a1[0] + str(d)
        sheet[cell] = contract[5]
        #status
        cell = a1[1] + str(d)
        sheet[cell] = contract[6]
        
    try:
        excelfile.save('.\\Excel Sheets\\Contracts.xlsx')
    except PermissionError:
        print('File is open')
        input('Press any key to continue')
        
def writenextsteps(nextsteps):  
    #write the next step data for inprocess tasks
    excelfile =  openpyxl.load_workbook('.\\Excel Sheets\\Contracts.xlsx')
    a = list(string.ascii_uppercase)
    d=2
    sheet2=excelfile['InProcess']
    for step in nextsteps:
        d=d+1
        #write states
        cell = a[0+7] + str(d)
        sheet2[cell] = step
        
    try:
        excelfile.save('.\\Excel Sheets\\Contracts.xlsx')
    except PermissionError:
        print('File is open')
        input('Press any key to continue')