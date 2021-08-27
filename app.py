from tkinter import font
from tkinter.constants import CENTER, DISABLED, HIDDEN, LEFT
import PySimpleGUI as sg
from email.mime import text
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
import webbrowser

sg.theme("dark grey 2")
layout = [
            [sg.Text("Welcome To Mail Automation By codeking5",font=20,text_color="skyblue",justification=CENTER,border_width=10,)],
            [sg.Text("Enter Your Email Address")],
            [sg.Input(tooltip="Enter Email Address",border_width=0)],
            [sg.Text("Enter Your Password")],
            [sg.Input(password_char="*",tooltip="Enter Email Password",border_width=0)],
            [sg.Text("Enter Mail Subject")],
            [sg.Input(tooltip="Enter Your Mail Subject",border_width=0)],
            [sg.Text("Select Your Email List csv file")],
            [sg.InputText(border_width=0), sg.FileBrowse(file_types=[("Email List CSV","*.csv")])],
            [sg.Text("Select Your Html Mail Template")],
            [sg.InputText(border_width=0), sg.FileBrowse(file_types=[("Select HTML Mail Template","*.html")])],
            [sg.Button("Send",size=(10,0),font=14)],
            [sg.Button("Made By codeking5",border_width=0)],

        ]
window = sg.Window('Mail Automation By codeking5', layout)
event, values = window.Read()
user_email = values[0]
user_password = values[1]
mailSubject = values[2]
csvFile = values[3]
htmlTemplete = values[4]

window.AutoClose = False
if(event == "Made By codeking5"):
    webbrowser.open("https://github.com/codeking5",new=1)
    event, values = window.Read()

if(event == "Send"):
    if(csvFile != None and htmlTemplete != None and user_email != None and user_password != None):
        email_list = pd.read_csv(csvFile)
        form_Address = user_email
        to_Address = email_list['Email']

        msg = MIMEMultipart()
        msg['From'] = form_Address
        msg['To'] =" ,".join(to_Address)
        msg['subject'] = mailSubject

        with open(htmlTemplete) as file:
            body = file.read()

        msg.attach(MIMEText(body,'html'))

        email=user_email
        password=user_password

        mail = smtplib.SMTP("smtp.gmail.com",587)
        mail.ehlo()
        mail.starttls()
        mail.login(email,password)
        text = msg.as_string()
        mail.sendmail(form_Address,to_Address,text)
        mail.quit()
        event, values = window.Read()


window.close()