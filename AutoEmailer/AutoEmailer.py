import win32com.client as win32
import os
import pandas as pd

outlook = win32.DispatchEx('outlook.application')   #TODO: modify to use any e-mail, not just control Outlook

#file paths
path_root = "C:\\Users\\skirkpatrick\\Documents\\Work for Art\\DG Pledge Notification Reports\\"
path_contactlist = "//concordia/lancentral/Work for Art/Designated Gifts/DG Pledge Reports//"   #network locations use forward slashes
path_input = "C:\\Users\\skirkpatrick\\Documents\\Work for Art\\DG Pledge Notification Reports\\Input\\"
path_output = "C:\\Users\\skirkpatrick\\Documents\\Work for Art\\DG Pledge Notification Reports\\Output\\"
path_errors = "C:\\Users\\skirkpatrick\\Documents\\Work for Art\\DG Pledge Notification Reports\\Errors\\"
file_emailbody = path_root + "email template.txt"
file_contactlist = path_contactlist + "DG Arts Org Email List for Pledge reports.xlsx"

#load data
message_body = open(file_emailbody).read()  #load message body from file
wb = pd.ExcelFile(file_contactlist)
df1 = wb.parse('Contacts')  #parse contents of excel file into Pandas dataframe

#initialize lists
files_good = []
files_error = []

#email loop
for f in os.listdir(path_input):
    _f = f.split(" - ")                 #split filename of form "Org Name - New Pledges through Work for Art.xlxs" into array _f
    validemailacount = 0
    emaillist = []      #list of e-mails to send to (each item might need to be mailed to multiple addresses)
    for index, row in df1.iterrows():   #iterate over all rows, searching for matches
        if row[1] == _f[0]:             #if Org Name matches File Name, add e-mail to list
            emaillist.append(row[6])
            validemailacount += 1
    if validemailacount == 0:    #if no valid matches, print error and save filename in error list
        files_error.append(f)
        print(_f[0], " not found!")
    else:   #if there is at least one matching e-mail, send e-mail
        stringemail = ""
        print("Org: ", _f[0])
        files_good.append(f)    #add filename to good list
        for address in emaillist:
            stringemail += address + ";"
        print("Sending email to: ", stringemail)
        mail = outlook.CreateItem(0)
        mail.To = stringemail
        mail.Subject = 'New Pledges through Work for Art'
        mail.Body = message_body
        mail.Attachments.Add(path_input + f) #f is the excel report filename
        mail.Send()

#file moving loops
for f in files_good:
    os.rename(path_input + f, path_output + f)
for f in files_error:
    os.rename(path_input + f, path_errors + f)