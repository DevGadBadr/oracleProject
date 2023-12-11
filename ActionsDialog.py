from actionsUI import Ui_Dialog
from PyQt5.QtWidgets import QApplication,QDialog,QWidget
from PyQt5 import QtWidgets , QtCore,QtGui
from PyQt5.QtCore import Qt,pyqtSignal,QThread
from functools import partial
from newoulook import Emailer
from pythoncom import com_error




class emailerr(QThread):
    finished = pyqtSignal()
    def __init__(self,row):
        super().__init__()
        self.row = row

    def run(self):
        print('Ok for email')
        x=Actionswindow()
        text = x.readtext()
        emailbodylist = text[self.row]
        status = emailbodylist[-1]
        
        if 'Project Operations Director' in status:
            person = 'basma.taha@madkour.com.eg'
        elif 'Infrastructure Sector Director' in status:
            person = 'basma.taha@madkour.com.eg'
        elif 'Infrastructure Manager Of Projects' in status:
            person = 'amin.youssry@madkour.com.eg'
        elif 'PM_Mahmoud Ibrahim' in status:
            person = 'mahmoud.ibrahim@madkour.com.eg'
        elif 'Projects Control GRP' in status:
            person = 'mohamed.s.shehata@madkour.com.eg'
        elif 'Accounts Payable GRP_CVL' in status:
            person = 'mona.afifi@madkour.com.eg'
        elif 'Procurement Section Head' in status:
            person = 'mohamed.kobisi@madkour.com.eg'
        else:
            person=''
            
        
        greeting = 'Dear Valued,<br>Greetings,<br>We need your kind approval on the following,<br><br>'
        emailbody = ''
        for item in emailbodylist:
            emailbody = emailbody + item +'<br>'

        
        print(emailbody,self.row)
        
        try:
            Emailer(greeting+emailbody,'',person)
        except com_error:
            print('An Email is Already Open')

class Actionswindow(Ui_Dialog):
    
    def localview(self,window):
        super().setupUi(window)
        self.myedit(window)
        self.emptybox()
        window.setWindowIcon(QtGui.QIcon('logo.jpg'))
        window.setWindowTitle('Take Actions')
        
    def itemadder(self,row=0):
        self.groupBox = QtWidgets.QGroupBox(self.scrollAreaWidgetContents)
        self.groupBox.setGeometry(QtCore.QRect(10, 20+row*190, 820, 201))
        self.groupBox.setObjectName("groupBox")
        self.textBrowser = QtWidgets.QTextBrowser(self.groupBox)
        self.textBrowser.setGeometry(QtCore.QRect(10, 10, 790, 171))
        self.textBrowser.setObjectName("textBrowser")
        self.buttonBox = QtWidgets.QDialogButtonBox(self.groupBox)
        self.buttonBox.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.buttonBox.setAutoFillBackground(False)
        self.buttonBox.setGeometry(QtCore.QRect(650, 110, 193, 28))
        self.Email = QtWidgets.QDialogButtonBox.StandardButton.Ok
        self.Dismiss = QtWidgets.QDialogButtonBox.StandardButton.No
        self.buttonBox.setStandardButtons(self.Email)
        self.n = self.buttonBox.button(self.Email)
        self.n.setText('Email')
        self.n.clicked.connect(partial(self.emailsender,row))        
        self.buttonBox.setCenterButtons(True)
        self.buttonBox.setObjectName("buttonBox")

        textlist = self.readtext()
        projectdetails = textlist[row]
        for line in projectdetails:
            self.textBrowser.append(line)
        
        self.scrollArea.setWidget(self.scrollAreaWidgetContents)
        n=(row+1)*200
        self.scrollAreaWidgetContents.setGeometry(QtCore.QRect(0, 0, 800, n))
        self.textnumber = 'text ' + str(row)
        self.emailbuttonnumber = 'email but ' + str(row)
        self.dismissbuttonnumber = 'dismiss but ' + str(row)     
        item =(row,self.textnumber,self.emailbuttonnumber,self.dismissbuttonnumber)

        return item
        
        
    def myedit(self,window):
        
        self.scrollArea = QtWidgets.QScrollArea(window)
        self.scrollArea.setGeometry(QtCore.QRect(70, 90, 841, 521))
        self.scrollArea.setWidgetResizable(True)
        self.scrollArea.setObjectName("scrollArea")
        self.scrollAreaWidgetContents = QtWidgets.QWidget()
        self.scrollAreaWidgetContents.setGeometry(QtCore.QRect(0, 0, 800, 2500))
        self.scrollAreaWidgetContents.setObjectName("scrollAreaWidgetContents")
        self.scrollArea.setWidgetResizable(True)
        self.pushButton.clicked.connect(window.close)
        self.groupBox.setStyleSheet("QGroupBox { border: none; }")
        _translate = QtCore.QCoreApplication.translate
        self.groupBox.setTitle(_translate("Dialog", ""))
        window.setWindowFlags(window.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        window.setWindowFlags(window.windowFlags() | Qt.WindowMinimizeButtonHint)
        window.setFixedSize(960,706)
    
    def editscroll(self):
        self.scrollAreaWidgetContents.setMinimumSize(self.scrollAreaWidgetContents.width(), self.scrollAreaWidgetContents.height())
    
    def itemshandler(self,n):
        
        self.myitemlist=[]
        for item in range(0,n):
            self.itemadded = self.itemadder(item)
            self.myitemlist.append(self.itemadded)
        self.editscroll()
        
    
    def readtext(self):

        textdata = open('textboxesdata.txt','r',encoding='utf-8')
        boxesdata = textdata.read()
        pros = boxesdata.split('*')
        neatpros=[]
        for pro in pros:
            neatpro = pro.split('\n')
            neatpros.append(neatpro)

        for pro in neatpros:
            for index,item in enumerate(pro):
                if item =='':
                    pro.pop(index)
        neatpros.pop()
        neatpros

        return neatpros


    def emailsender(self,row):
        self.emilthread = emailerr(row)
        self.emilthread.start()
    
    def emptybox(self):
        font2 = QtGui.QFont()
        font2.setFamily("Segoe UI Black")
        font2.setPointSize(12)
        font2.setBold(True)
        font2.setWeight(60)
        self.label.setFont(font2)
        wid = QtWidgets.QWidget()
        label = QtWidgets.QLabel(wid)
        label.setText('\n\n\n\n\n\n\n\n\n \t\tThere is nothing in here Enjoy your Empty List')
        label.setFont(font2)
        self.scrollArea.setWidget(wid)


if __name__=='__main__':
    app= QApplication([])
    window = QDialog()
    
    x = Actionswindow()
    x.localview(window)
    window.show()
    app.exec()
    