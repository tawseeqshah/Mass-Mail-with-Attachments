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
    "D:\\Atom\\Python\\Projects Python\\Email Sending\\Student List.xlsx")
mail_column = mails.get("Email")
listmails = list(mail_column)
newlist = [x for x in listmails if pd.isnull(x) == False]
print(f"Sending mass mail to {newlist}.\n")

#Topic 1
mails = pd.read_excel(
    "D:\\Atom\\Python\\Projects Python\\Email Sending\\Student List.xlsx")
top1_column = mails.get("Topic 1")
listtop1 = list(top1_column)
top1list = [x for x in listtop1 if pd.isnull(x) == False]
print(top1list)

#Topic 2
mails = pd.read_excel(
   "D:\\Atom\\Python\\Projects Python\\Email Sending\\Student List.xlsx")
top2_column = mails.get("Topic 2")
listtop2 = list(top2_column)
top2list = [x for x in listtop2 if pd.isnull(x) == False]
print(top2list)

#NAmes
mails = pd.read_excel(
   "D:\\Atom\\Python\\Projects Python\\Email Sending\\Student List.xlsx")
name_column= mails.get("Name")
namelist = list(name_column)
name1list = [x for x in namelist if pd.isnull(x) == False]
print(name1list)

# Designation
mails = pd.read_excel(
    "D:\\Atom\\Python\\Projects Python\\Email Sending\\Student List.xlsx")
Designation_column= mails.get("Designation")
Designationlist = list(Designation_column)
Designation1list = [x for x in Designationlist if pd.isnull(x) == False]
print(Designation1list)

for i in range (len(top1list)):
    server = sm.SMTP("smtp.gmail.com", port=587)

    server.starttls()
    server.login("tawseeq_33btech20@nitsri.net", "Ta@taw32seeq01")

    from_ = "tawseeq_33btech20@nitsri.net"
    to_ = newlist[i]
    message = MIMEMultipart("alternative")
    message['X-Priority'] = '1'
    message["Subject"] = "Exploring research internship opportunity for Winter 2022-2023"
    message["From"] = "tawseeq_33btech20@nitsri.net"
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

    <p> I am writing this email in pursuit of an opportunity to work as a
    <b>research intern</b> under your eminent guidance for the upcoming <b>winter of
    2023</b>.<p/>

    <p/>I am veritably interested in mechanical engineering, especially
    pertaining to <b>{top1list[i]}</b>. I've been scrutinizing <b>{top2list[i]}</b> thoroughly. I have also read several
    publications of yours.<p/>

    <p>Recently, I went through your conspicuous work and  I was really
    intrigued by your participation and benefaction in the field. If
    possible, I would love to work under your esteemed supervision and am
    willing to work on any project that will help me fathom the dynamics
    of materials.<p/>

    <p>Please find my <b>CV attached</b> with this mail for your perusal.
    I look forward to hearing from you about my candidacy.<p/>

    <p>Thank you for your time and consideration.<p/>
    <p>Sincerely,<p/>
    <p>Tawseeq Mushtaq Shah<p/>
            
    </body>
    </html>
    '''

    text = MIMEText(html, "html")
    message.attach(text)
    attach_file_name = "D:\Atom\Python\Projects Python\Email Sending\Tawseeq's Resume (4).pdf"
    attach_file = open(attach_file_name, 'rb') # Open the file as binary mode

    payload = MIMEBase('application', 'octet-stream')
    payload.set_payload((attach_file).read())
    encoders.encode_base64(payload) #encode the attachment
    #add payload header with filename
    payload.add_header('Content-Disposition', 'attachment; filename="Resume.pdf"')
    message.attach(payload)
    server.sendmail(from_, to_, message.as_string())
    print(f"Mail has been sent to {Designation1list[i]} {name1list[i]}")
