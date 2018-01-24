import win32com.client as win32
import os
import pandas as pd

outlook = win32.DispatchEx('outlook.application')

filepath1 = "C:\\Users\\skirkpatrick\\Documents\\Work for Art\\DG Pledge Reports\\"
filepath2 = "1-16-2017\\"
filename_body = filepath1 + "email template.txt"
filename_contactlist = filepath1 + "Email List.xlsx"

message_body = open(filename_body).read()

wb = pd.ExcelFile(filename_contactlist)
df1 = wb.parse('Contacts')
bademaillist = []   #list of orgs with missing/bad contact info

for f in os.listdir(filepath1 + filepath2):
    _f = f.split(" - ") #split filename
    #print("ORG NAME: ", _f[0])        #grab first portion of name, use for comparing to Organization Name
    validemailacount = 0
    emaillist = []      #list of e-mails to send to
    for index, row in df1.iterrows():   #iterate over all rows, searching for matches
        if row[1] == _f[0]:             #if Org Name matches File Name, add e-mail to list
            #print("e-mail: ", row[6])
            emaillist.append(row[6])
            validemailacount += 1
    if validemailacount == 0:    #if there are No valid matches, throw exception
        bademaillist.append(_f[0])
        print(_f[0], " not found!")
    else:   #if there is at least one matching e-mail, send e-mail
        stringemail = ""
        print("Org: ", _f[0])
        for address in emaillist:
            stringemail += address + ";"
        print("Sending email to: ", stringemail)
        mail = outlook.CreateItem(0)
        mail.To = stringemail
        mail.Subject = 'New Pledges through Work for Art'
        mail.Body = message_body
        mail.Attachments.Add(filepath1 + filepath2 + f) #f is the excel report filename
        mail.Send()

print("Orgs not found: ", bademaillist)