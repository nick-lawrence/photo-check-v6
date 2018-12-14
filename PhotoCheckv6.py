#!/Users/adassist/anaconda3/bin/python
#PhotoCheck v6: 'The Burner'

print("Loading PhotoCheck v6...\n")

import os
import datetime
from shutil import copyfile
import xlrd
import re
import webbrowser
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from fpdf import FPDF
from private import password, server-address

#Today's date as string variable 'mm-dd-yy'
date = datetime.datetime.today().strftime('%m-%d-%Y') 

#Sold List Origin // server location variable
SLO = '/Volumes/255CMS/RICHARD/Sold list.xls'

#Copied Sold List // workstation location variable
CSL = '/Users/adassist/Documents/PhotoCheck v6/Sold List for Checking/Sold list.xls'

#Copy Sold List from server to workstation for read/write
print("Copying Files...\n")
copyfile(SLO, CSL)

#Open copied sold list as 'workbook' variable
workbook = xlrd.open_workbook(CSL)
worksheet = workbook.sheet_by_index(0)

#Set forms_list to nothing
forms_list = []

#Open worksheet and add all rows to forms_list
for x in range(0,500):
    forms_list.append(worksheet.row(x))

#Flatten forms_list, get rid of empties and dates
flat_forms_list = []
for sublist in forms_list:
    for item in sublist:
        if 'empty' not in str(item) and '-' in str(item):
            flat_forms_list.append(str(item))

print("Collating List...\n")

#Get rid of all extraneous characters and text in list items\
contains_wcml = False
if "WCML" in str(flat_forms_list):
    contains_wcml = True

new = str(flat_forms_list).replace('text:', '')
new = str(new).replace("'", '')
new = str(new).replace("'", '')
new = str(new).replace("[", '')
new = str(new).replace("]", '')
new = str(new).replace('"', '')
new = str(new).replace(' ', '')
new = str(new).replace('ML', 'WCML')

#Get rid of the two uneeded front characters of each item '1-', '2-' etc.
newer = []
new = new.split(',')
for item in new:
    if item[0].isnumeric():
        cos = item.lstrip(item[0])
        cos = cos.lstrip(cos[0])
        newer.append(cos)
    else:
        newer.append(item)

#Open items in web browser tabs for manual review
print("Opening Browswer Tabs for Manual Review...\n")
for item in newer:
    webbrowser.open(f"https://rmi-online.com/{item.lower()}.html")


#Set lists to nothing and prompt user for manual review decision
photos_needed = []
error_pages_404 = []

for item in newer:
    choice_loop = True
    while choice_loop == True:
        choice = input(f"Please input code for form {item}:\n(0 = Not Needed, 1 = Photo Needed, 2 = 404 Error Page)\n")
        if choice == '1':
            photos_needed.append(item)
            break
        elif choice == '2':
            error_pages_404.append(item)
            break
        elif choice == '0':
            break
        else:
            print("Invalid option, please try again.")

print("List Complete, loading Email...\n")

#Send email to McAlpine with list of 404 error pages
print("Contacting Server...")

msg = MIMEMultipart()

if contains_wcml == True:
    message = f"Hey Mark, here are the 404 error items for {date}: \n WARNING: This list contains a WCML! \n " + ("\n").join(error_pages_404)
else:
    message = f"Hey Mark, here are the 404 error items for {date}: \n" + ("\n").join(error_pages_404)
 
# setup the parameters of the message
password = password()
msg['From'] = "nick@rmi-online.com"
msg['To'] = "mcalpine@rmi-online.com"
msg['Subject'] = f"404 Error Photos | {date}"
 
# add in the message body
msg.attach(MIMEText(message, 'plain'))
 
#create server
server = smtplib.SMTP(server_address())
 
server.starttls()
 
# Login Credentials for sending the mail
server.login(msg['From'], password)
 
 
# send the message via the server.
server.sendmail(msg['From'], msg['To'], msg.as_string())
 
server.quit()
 
print("Successfully sent email to %s:" % (msg['To']))

#######################################################
#### DAY 2: BUILD THE PDF RENDERER FOR PRODUCTION #####
#######################################################
 
#Generate PDF saved in same directory, called 'Photos Needed {date}'.pdf 
print("Rendering PDF...\n")
#Copy 1 of photos_needed for production
pdf = FPDF()
pdf.add_page()
pdf.set_font("Arial", size=30)
pdf.cell(200, 10, txt=f"PHOTOS NEEDED {date}", ln=1, align="L")
pdf.cell(200, 10, txt="\n", ln=1, align="L")
for item in photos_needed:
    pdf.set_font("Arial", size=20)
    pdf.cell(200, 10, txt=f"{item}", ln=1, align="L")
pdf.cell(200, 10, txt="\n", ln=1, align="L")
pdf.cell(200, 10, txt="\n", ln=1, align="L")
pdf.cell(200, 10, txt="\n", ln=1, align="L")
pdf.cell(200, 10, txt="\n", ln=1, align="L")

#Copy 2 of photos_needed for advertising
pdf.set_font("Arial", size=30)
pdf.cell(200, 10, txt=f"PHOTOS NEEDED {date}", ln=1, align="L")
pdf.cell(200, 10, txt="\n", ln=1, align="L")
for item in photos_needed:
    pdf.set_font("Arial", size=20)
    pdf.cell(200, 10, txt=f"{item}", ln=1, align="L")
pdf.cell(200, 10, txt="\n", ln=1, align="L")
pdf.cell(200, 10, txt="\n", ln=1, align="L")
pdf.cell(200, 10, txt="\n", ln=1, align="L")
pdf.cell(200, 10, txt="\n", ln=1, align="L")

#Render PDF
pdf_name = f"Photos Needed {date}.pdf"
pdf.output('/Users/adassist/Desktop/Photo_Check v6/Photos Needed/'+pdf_name)

print("Rendering Complete...\n")
print("Execution Complete.\n")

#Open PDF
### Issue fixed, the PDF should open now.
### Note: MAKE ABSOLUTELY SURE THAT OS.SYSTEM is passed a *shell command* (like Open), NOT JUST A FILEPATH!!!
os.system("open '/Users/adassist/Desktop/Photo_Check v6/Photos Needed/Photos Needed '"+ date +"'.pdf'")
