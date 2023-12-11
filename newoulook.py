
import win32com.client as win32   
import win32com
import pythoncom


def Emailer(text, subject, recipient, auto=True):

    xl=win32com.client.Dispatch("outlook.application",pythoncom.CoInitialize())

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = subject
    mail.HtmlBody = text
    if auto:
        mail.Display(True)
    else:
        mail.open # or whatever the correct code is


