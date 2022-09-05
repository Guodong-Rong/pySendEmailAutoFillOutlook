#requirements.txt add for py 3 -> pypiwin32
import win32com.client as win32
import os
import pandas as pd
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from contextlib import redirect_stdout
import smtplib
import xlwings as xw

# from mailer import Mailer

# Read the file
workbook = xw.Book('LeaderAccountExcel.xlsx')
email_list = workbook.sheets[0].range('A1').options(pd.DataFrame, 
                                            header=1,
                                            index=False, 
                                            expand='table').value
# Replace None as blank string
email_list.fillna("",inplace=True)
# email_list = pd.read_excel("LeaderAccountExcel.xlsx", engine="openpyxl")

all_fullName = email_list['Full Name']
all_firstName = email_list['First Name']
all_surname = email_list['Surname']
all_phoneNumber = email_list['Phone Number']
all_userName = email_list['User Name']
all_password = email_list['Password']
all_companyName = email_list['Company Name']
all_subject = email_list['Subject']
all_pictures = email_list['Picture']
all_attachments = email_list['Attachment']
all_body = email_list['Body']

# your_name = "Key Institute"
# your_email = "sylvianron@outlook.com"
# your_password = "609716230zwj"
def Emailer(fullName, password, companyName, text, subject, recipient, logo, attachment):
    

    outlook = win32.Dispatch('outlook.application')
    
    # mail = Mailer(email='admin@keyinstitute.com.au', password='Training2020!')
    From = None
    for myEmailAddress in outlook.Session.Accounts:
        if "admin@keyinstitute.com.au" in str(myEmailAddress):
            From = myEmailAddress
            break

    mail = outlook.CreateItem(0)
    mail._oleobj_.Invoke(*(64209, 0, 8, 0, From))
    
    mail.To = recipient
    mail.Subject = subject
    
    ###
    # outer = MIMEMultipart()

    # outer.attach(MIMEText(text, 'html')) # or 'html'/'plain'
    # outer.preamble = 'You will not see this in a MIME-aware mail reader.\n'
    # picture = logo
    # #Attach Image 
    # fp = open(picture, 'rb') #Read image 
    # print(str(picture))
    # msgImage = MIMEImage(fp.read())
    # fp.close()

    # Define the image's ID as referenced above
    # msgImage.add_header('Content-ID', '<logo>')
    # msgImage.add_header('Content-Disposition', 'inline', filename='filename')
    # outer.attach(msgImage)
    # attachment = mail.Attachments.Add(picture)
    # attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "logo")
    # try:
    #     if (all_attachments[idx] == all_attachments[idx]):
    #         for file in attachments:
    #             with open(file, 'rb') as fp:
    #                 msg = MIMEBase('application', "octet-stream")
    #                 msg.set_payload(fp.read())
    #             encoders.encode_base64(msg)
    #             msg.add_header('Content-Disposition','attachment', filename=os.path.basename(file))
    #             outer.attach(msg)
    #     except:
    #         print('Unable to open one of the attachments.', sys.exc_info()[0],'\n\n')
    # mail.Attachments.Add(picture)
    mail.HtmlBody = text

    # attachment1 = os.getcwd() +"\\file.ini"
    # print(attachment)
    if attachment is not "":
        attachmentArray = attachment.split(";")
        for i in range(len(attachmentArray)):
            mail.Attachments.Add(attachmentArray[i])
    # mail.Attachments.Add(attachment)

    ###
    mail.Display(True)

# MailSubject= "Auto test mail"
# # MailInput="""
# # #html code here
# # """
# MailAdress="gradon.rong@gmail.com"

for idx in range(len(all_userName)):
    
    # all_body[idx] = str(all_body[idx]).replace("${firstName}",str(all_firstName[idx]))
    # all_body[idx] = str(all_body[idx]).replace("${surname}",str(all_surname[idx]))
    # all_body[idx] = str(all_body[idx]).replace("${phoneNumber}",str(all_phoneNumber[idx]))
    # all_body[idx] = str(all_body[idx]).replace("${userName}",str(all_userName[idx]))
    # all_body[idx] = str(all_body[idx]).replace("${password}",str(all_password[idx]))
    # all_body[idx] = str(all_body[idx]).replace("${companyName}",str(all_companyName[idx]))
    for r in (("${FullName}",str(all_fullName[idx])),
    ("${UserName}",str(all_userName[idx])),
    ("${Password}",str(all_password[idx])),
    ("${CompanyName}",str(all_companyName[idx]))):

        all_body[idx] = str(all_body[idx]).replace(*r)
    
    Emailer(all_fullName[idx], all_password[idx], all_companyName[idx], all_body[idx], all_subject[idx], all_userName[idx], all_pictures[idx], all_attachments[idx]) #that open a new outlook mail even outlook closed.

