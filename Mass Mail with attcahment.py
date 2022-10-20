from email.mime import text
from fileinput import filename
from numpy import string_
import pandas as pd
import smtplib as sm
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders
from email.mime.base import MIMEBase

#Mails list print.......................................................................................

mails = pd.read_excel(
    "File_path")
mail_column = mails.get("Email")
listmails = list(mail_column)
newlist = [x for x in listmails if pd.isnull(x) == False]
print(f"Sending mass mail to {newlist}.\n")

#Topic 1
mails = pd.read_excel(
    "File_path")
top1_column = mails.get("Topic 1")
listtop1 = list(top1_column)
top1list = [x for x in listtop1 if pd.isnull(x) == False]
print(top1list)

#Topic 2
mails = pd.read_excel(
   "File_path")
top2_column = mails.get("Topic 2")
listtop2 = list(top2_column)
top2list = [x for x in listtop2 if pd.isnull(x) == False]
print(top2list)

#NAmes
mails = pd.read_excel(
   "File_path")
name_column= mails.get("Name")
namelist = list(name_column)
name1list = [x for x in namelist if pd.isnull(x) == False]
print(name1list)

# Designation
mails = pd.read_excel(
    "File_path")
Designation_column= mails.get("Designation")
Designationlist = list(Designation_column)
Designation1list = [x for x in Designationlist if pd.isnull(x) == False]
print(Designation1list)

for i in range (len(top1list)):
    server = sm.SMTP("smtp.gmail.com", port=587)

    server.starttls()
    server.login("from_mail_id", "password")

    from_ = "from_mail_id"
    to_ = newlist[i]
    message = MIMEMultipart("alternative")
    message['X-Priority'] = '1'
    message["Subject"] = "My subject"
    message["From"] = "from_mail_id"
    print(f"Sending {i+1}st/th mail to {newlist[i]}.\n")

    html = f'''

    <html>
    <head>

    </head>
    <body>
    <p>Dear {Designation1list[i]} {name1list[i]},</p>

    <p>Hope this email finds you well. I am <b>Tawseeq Mushtaq Shah</b>, a 3rd-year
    (5th semester) Mechanical Engineering undergraduate at the <b>National
    Institute of Technology Srinagar</b>.<p/>
    
    <p/>I am veritably interested in mechanical engineering, especially
    interested in <b>{top1list[i]}</b> and <b>{top2list[i]}</b> thoroughly.<p/>

    <p>Please find my <b>CV attached</b> with this mail for your perusal.<p/>
    
    <p>Thank you for your time and consideration.<p/>
    <p>Sincerely,<p/>
    <p>Tawseeq Mushtaq Shah<p/>
            
    </body>
    </html>
    '''

    text = MIMEText(html, "html")
    message.attach(text)
    attach_file_name = "attachment.pdf"
    attach_file = open(attach_file_name, 'rb') # Open the file as binary mode

    payload = MIMEBase('application', 'octet-stream')
    payload.set_payload((attach_file).read())
    encoders.encode_base64(payload) #encode the attachment
    #add payload header with filename
    payload.add_header('Content-Disposition', 'attachment; filename="Resume.pdf"')
    message.attach(payload)
    server.sendmail(from_, to_, message.as_string())
    print(f"Mail has been sent to {Designation1list[i]} {name1list[i]}")
