import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import pandas as pd
import csv

class student(object):
    def __init__(self,roll,name,email_id) -> None:
        self.roll = roll
        self.name = name
        self.email_id = email_id
        self.file=None

    def __repr__(self):
        return f"{self.roll} {self.name} {self.email_id} {self.file}"

class document(object):
    def __init__(self,name,path) -> None:
        self.name = name
        self.path = path
    def __repr__(self) -> str:
        return f"{self.name} {self.path}"

def send_email(sender_address,receiver_address, sender_email_pass, email_subject, email_body, email_attachment):
    # Create the email
    msg = MIMEMultipart()
    msg['From'] = sender_address
    msg['To'] = receiver_address
    msg['Subject'] = email_subject
    msg.attach(MIMEText(email_body, 'plain'))

    # Attach the attachment
    filename = email_attachment
    file_address = os.path.join(os.getcwd(), filename)
    attachment = open(file_address, "rb")
    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment", filename = filename)
    msg.attach(part)

    # Send the email
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(sender_address, sender_email_pass)
    text = msg.as_string()
    server.sendmail(sender_address, receiver_address, text)
    server.quit()
    print(f'Mail Sent to {receiver_address}')

def convert_extract(xlsxFile,student_list):

    read_file = pd.read_excel(xlsxFile)
    read_file.to_csv('test.csv', index=None,header=True)
    data = pd.read_csv('test.csv',usecols=['Name','Roll No','Mail Id'])
    data.to_csv('test.csv', index=None,header=True)

    with open('test.csv', mode ='r')as file:
        # reading the CSV file
        csvFile = csv.reader(file)
        # displaying the contents of the CSV file
        for lines in csvFile:
            student_list.append(student(lines[1],lines[0],lines[2]))

def docs(document_list):
    path = os.getcwd()
    DirList= os.listdir(path)
    for i in range(len(DirList)):
        document_list.append(document(os.path.basename(DirList[i]),f'{os.path.join(os.getcwd(), DirList[i])}'))
    return DirList

def linkFiles(doc,stu):
    for f in doc:
        for student in stu:
            if str(student.roll) in str(f.name):
                # print(str(student.roll),f.name)
                student.file = str(f.name)

def sendToAll(student_list):
    for student in student_list[1:]:
        if student.file:
            send_email('agrawalsm_4@rknec.edu',str(student.email_id),'Legendary@','hello','hello',str(student.file))

# crearing list of documents     
document_list=[]
docs(document_list)

# Student info with roll name and email
student_list=[]
# creating student list having objects of student class
convert_extract('internals.xlsx',student_list)

linkFiles(document_list,student_list)
# link the paths with the student list


sendToAll(student_list)