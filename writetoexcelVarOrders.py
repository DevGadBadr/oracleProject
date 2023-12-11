#Preparing Excel File indices            
import openpyxl
import string

def writetoexcelfn(inprocess,rejected,approved,closed,Draft):
    
    a = list(string.ascii_uppercase)
    a1 = []
    for n in range(26):
        a1.append('A' + a[n])

    excelfile =  openpyxl.load_workbook('.\\Excel Sheets\\Variation Orders.xlsx')
    sheet= excelfile['Sheet1']
    
    #Deleting Previous Data Sheet1
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
        #write payment
        cell = a[0] + str(d)
        sheet[cell] = contract[0]
        #write docnumber
        cell = a[1] + str(d)
        sheet[cell] = contract[1]
        #project
        cell = a[2] + str(d)
        sheet[cell] = contract[2]
        #subcontractor 
        cell = a[3] + str(d)
        sheet[cell] = contract[3]
        #subcontractor 
        cell = a[4] + str(d)
        sheet[cell] = contract[4]
        #total 
        cell = a[5] + str(d)
        sheet[cell] = contract[5]
        #date
        cell = a[6] + str(d)
        sheet[cell] = contract[6]
        #status
        cell = a[7] + str(d)
        sheet[cell] = contract[7]

    #Approved
    d=2
    for contract in approved:
        d=d+1
        #write payment
        cell = a[0+8] + str(d)
        sheet[cell] = contract[0]
        #write docnumber
        cell = a[1+8] + str(d)
        sheet[cell] = contract[1]
        #project
        cell = a[2+8] + str(d)
        sheet[cell] = contract[2]
        #subcontractor 
        cell = a[3+8] + str(d)
        sheet[cell] = contract[3]
        #subcontractor 
        cell = a[4+8] + str(d)
        sheet[cell] = contract[4]
        #total 
        cell = a[5+8] + str(d)
        sheet[cell] = contract[5]
        #date
        cell = a[6+8] + str(d)
        sheet[cell] = contract[6]
        #status
        cell = a[7+8] + str(d)
        sheet[cell] = contract[7]

    #Rejected
    d=2
    for contract in rejected:
        d=d+1
        #write payment
        cell = a[0+16] + str(d)
        sheet[cell] = contract[0]
        #write docnumber
        cell = a[1+16] + str(d)
        sheet[cell] = contract[1]
        #project
        cell = a[2+16] + str(d)
        sheet[cell] = contract[2]
        #subcontractor 
        cell = a[3+16] + str(d)
        sheet[cell] = contract[3]
        #subcontractor 
        cell = a[4+16] + str(d)
        sheet[cell] = contract[4]
        #total 
        cell = a[5+16] + str(d)
        sheet[cell] = contract[5]
        #date
        cell = a[6+16] + str(d)
        sheet[cell] = contract[6]
        #status
        cell = a[7+16] + str(d)
        sheet[cell] = contract[7]

    #Closed
    d=2
    for contract in closed:
        d=d+1
        #write payment
        cell = a[0+24] + str(d)
        sheet[cell] = contract[0]
        #write docnumber
        cell = a[1+24] + str(d)
        sheet[cell] = contract[1]
        #project
        cell = a1[0] + str(d)
        sheet[cell] = contract[2]
        #subcontractor 
        cell = a1[1] + str(d)
        sheet[cell] = contract[3]
        #subcontractor 
        cell = a1[2] + str(d)
        sheet[cell] = contract[4]
        #total 
        cell = a1[3] + str(d)
        sheet[cell] = contract[5]
        #date
        cell = a1[4] + str(d)
        sheet[cell] = contract[6]
        #status
        cell = a1[5] + str(d)
        sheet[cell] = contract[7]
        
    #Draft
    d=2
    for contract in Draft:
        d=d+1
        #write payment
        cell = a1[6] + str(d)
        sheet[cell] = contract[0]
        #write docnumber
        cell = a1[7] + str(d)
        sheet[cell] = contract[1]
        #project
        cell = a1[8] + str(d)
        sheet[cell] = contract[2]
        #subcontractor 
        cell = a1[9] + str(d)
        sheet[cell] = contract[3]
        #subcontractor 
        cell = a1[10] + str(d)
        sheet[cell] = contract[4]
        #total 
        cell = a1[11] + str(d)
        sheet[cell] = contract[5]
        #date
        cell = a1[12] + str(d)
        sheet[cell] = contract[6]
        #status
        cell = a1[13] + str(d)
        sheet[cell] = contract[7]

    
    try:
        excelfile.save('.\\Excel Sheets\\Variation Orders.xlsx')
    except PermissionError:
        print('File is open')
        input('Press any key to continue')


def writenextsteps(inprocess,nextsteps):  
    #write the next step data for inprocess tasks
    excelfile =  openpyxl.load_workbook('.\\Excel Sheets\\Variation Orders.xlsx')
    a = list(string.ascii_uppercase)
    
    sheet2=excelfile['InProcess']
    
    #InProcess
    d=2
    for contract in inprocess:
        d=d+1
        #write payment
        cell = a[0] + str(d)
        sheet2[cell] = contract[0]
        #write docnumber
        cell = a[1] + str(d)
        sheet2[cell] = contract[1]
        #project
        cell = a[2] + str(d)
        sheet2[cell] = contract[2]
        #subcontractor 
        cell = a[3] + str(d)
        sheet2[cell] = contract[3]
        #subcontractor 
        cell = a[4] + str(d)
        sheet2[cell] = contract[4]
        #total 
        cell = a[5] + str(d)
        sheet2[cell] = contract[5]
        #date
        cell = a[6] + str(d)
        sheet2[cell] = contract[6]
        #status
        cell = a[7] + str(d)
        sheet2[cell] = contract[7]
    d=2
    for step in nextsteps:
        d=d+1
        #write states
        cell = a[0+8] + str(d)
        sheet2[cell] = step
        
    try:
        excelfile.save('.\\Excel Sheets\\Variation Orders.xlsx')
    except PermissionError:
        print('File is open')
        input('Press any key to continue')
        