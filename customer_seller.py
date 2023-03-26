import smtplib
import re
from openpyxl import load_workbook
from email.mime.text import MIMEText
import asyncio


mails = []

wb = load_workbook(filename = 'base.xlsx')
ws = wb.active

for row in ws.iter_rows(min_row=5, min_col=8, max_col=8, max_row=8, values_only=True):
    for iter_row in row:
        if iter_row != None:
            str(iter_row)
            mails.append(iter_row)
            # print(iter_row)
    

def send_email(message, receiver):
    sender = "lomex2014@mail.ru"
    # your password = "your password"
    password = "aPLznLee2G7TiaMLcRxf"
    # brugrmgulsghtwok
    server = smtplib.SMTP_SSL("smtp.mail.ru", 465)
    print(receiver)
 
    try:
        server.login(sender, password)
        msg = MIMEText(message)
        msg["Subject"] = "Это проверочное письмо"
        server.sendmail(sender, receiver, msg.as_string())

        # server.sendmail(sender, sender, f"Subject: CLICK ME PLEASE!\n{message}")

        return "The message was sent successfully!"
    except Exception as _ex:
        return f"{_ex}\nCheck your login or password please!"

print(len(mails))



for mail in mails:
        message = "Проверочное письмо"
        print(send_email(message=message, receiver=str(mail)))
        
        


