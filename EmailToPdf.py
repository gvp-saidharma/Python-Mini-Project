# Importing libraries
import imaplib, email
import cryptocode
import stdiomask
import os
from fpdf import FPDF
import random  

imap_url = 'imap.gmail.com'
user = 'isecuretester@gmail.com'
password = stdiomask.getpass()
# input('Enter your password:')

# Function to get email content part i.e its body part
def get_body(msg):
    if msg.is_multipart():
        return get_body(msg.get_payload(0))
    else:
        return msg.get_payload(None, True)
 
# Function to search for a key value pair
def search(key, value, con):
    result, data = con.search(None, key, '"{}"'.format(value))
    return data

def createDir(detach_dir, folder_name):
    detach_dir = '.'
    if folder_name not in os.listdir(detach_dir):
        os.mkdir(folder_name)
 
# Function to get the list of emails under this label
def get_email_pdf(result_bytes):
    msgs = [] # all the email data are pushed inside an array
    detach_dir = '.'
    folder_name = 'attachments'
    filePath = ""
    fileName = ""
    createDir(detach_dir, 'mails')
    createDir(detach_dir, folder_name)

    # Iterating over all emails
    for msgId in result_bytes[0].split():
        attachemts = []
        typ, messageParts = con.fetch(msgId, '(RFC822)')
        if typ != 'OK':
            print('Error fetching mail.')
            raise

        emailBody = messageParts[0][1]
        # print(emailBody)
        mail = email.message_from_bytes(emailBody)

        subject = mail.get("SUBJECT")
        from_mail = mail.get("FROM")
        to_mail = mail.get("TO")
        date = mail.get('DATE')
        # msgs.append('subject:'+ subject+ ",from_mail:"+ from_mail+ ",to_mail:"+ to_mail+ ",date:"+ date)

        for part in mail.walk():
            if part.get_content_maintype() == 'multipart':
                # print(part.as_string())
                continue
            if part.get('Content-Disposition') is None:
                # print('Content-Disposition')
                # print(part.as_string())
                # print(part.get_payload())
                msgs.append(part.get_payload())
                continue
            fileName = part.get_filename()

            if bool(fileName):           
                filePath = os.path.join(detach_dir, 'attachments', fileName)
                attachemts.append([fileName, filePath])
                if not os.path.isfile(filePath) :
                    msgs.append('fileName:'+ fileName)
                    fp = open(filePath, 'wb')
                    fp.write(part.get_payload(decode=True))
                    fp.close()
        
        body = get_body(mail)
        body = body.decode('UTF-8')
        # print(body)
        pdf = FPDF()

        pdf.add_page()
        pdf.set_font("Arial", size = 12)
        pdf.cell(5,10,txt="DATE "+date, ln=True)
        pdf.cell(5,10,txt="FROM "+from_mail, ln=True)
        pdf.cell(150,10,txt="TO  "+to_mail, ln=1)
        pdf.cell(150,5,txt="Subject :  "+subject, ln=1)
        pdf.cell(5,5,txt='',ln=True)

        if bool(fileName):
            for att in attachemts:
                pdf.cell(5,10,txt="Atachhment : " + att[0],ln=True,link= os.path.abspath(att[1]))
            pdf.multi_cell(0, 5, txt=body)
            pdf.cell(5,5,txt='',ln=True) 
        else:
            pdf.multi_cell(0, 5, txt=body)
            pdf.cell(5,5,txt='',ln=True)

        # pdf.output("./mails/" + subject + "_" + str(random.randint(100, 200)) + ".pdf")
        pdf.output("./mails/" + subject + ".pdf")



# this is done to make SSL connection with GMAIL
con = imaplib.IMAP4_SSL(imap_url)
 
# logging the user in
con.login(user, password)
 
# calling function to check for email under this label
con.select('Inbox')
 
# fetching emails from this user "saidharmareddy@gmail.com"
data = search('FROM', 'saidharmareddy@gmail.com', con)
get_email_pdf(data)
