import os
from pyairtable import Api
import jinja2
import pdfkit
from email.message import EmailMessage
from email import encoders
import ssl
import smtplib
from email.mime.application import MIMEApplication


#Connect Airtable
api = Api(os.environ['AIRTABLE_API'])

#Connect to Base and Table respectively
table = api.table("appJghtvtR8kehC1T", "tblGAv1OHRXpcc0eT")

#Load template into python
template_loader = jinja2.FileSystemLoader("./")
template_env = jinja2.Environment(loader=template_loader)    

def file_exists(rows):
    #This function checks if there is a file already created for the row in Airtable
    return os.path.exists(f'{rows["fields"]["ID"]}.pdf')

def print_to_pdf(rows):
    #This function prints content to PDF
    full_name = rows["fields"]["Full Name"]
    citizenship = rows["fields"]["Citizenship"]
    residence = rows["fields"]["Residence"]
    email_address = rows["fields"]["Email Address"]
    wallet_address = rows["fields"]["Wallet Address"]
    context = {
        "full_name": full_name, 
        "citizenship":citizenship, 
        "residence": residence, 
        "email_address": email_address, 
        "wallet_address": wallet_address
        }
    
    #Print PDF
    template = template_env.get_template("template.html")
    output_text = template.render(context)
    config = pdfkit.configuration(wkhtmltopdf="/usr/local/bin/wkhtmltopdf")
    pdfkit.from_string(output_text, f"{rows['fields']['ID']}.pdf", configuration=config)    

def email_pdf(rows):
    #This function sends a mail with a specified title body and attachment for each row
    email_sender = "grikunduz@gmail.com"
    email_password = os.environ.get("GMAIL_APP_PD")
    email_receiver = "justin@ortege.io"

    subject = "ORtege User Aggrement PDF"
    mail_body =  """
        Hey There is a new user just signed up. Here is the aggreement. See the attached pdf file.
    """

    em = EmailMessage()

    em ["From"] = email_sender
    em["To"] = email_receiver
    em["Subject"] = subject
    em.set_content(mail_body)

    #Make the message multipart
    em.add_alternative(mail_body, subtype="html")

    #Attach the pdf
    with open(f'{rows["fields"]["ID"]}.pdf', "rb") as attachment_file:
        file_data = attachment_file.read()
        file_name = f'{rows["fields"]["ID"]}.pdf'  

    #Create an instance of MIMEApplication and encode the attachment
    attachment = MIMEApplication(file_data, Name=file_name)
    attachment.set_payload(file_data)
    encoders.encode_base64(attachment)
    attachment.add_header('Content-Disposition', 'attachment', filename=file_name)
    em.attach(attachment)

    #Add SSL
    context = ssl.create_default_context()

    #Login and send the email
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as smtp:
        smtp.login(email_sender, email_password)
        smtp.sendmail(email_sender, email_receiver, em.as_string())
    
#iterate through each row and create a pdf file.
for rows in table.all():
    if not file_exists(rows):
        print_to_pdf(rows)
        email_pdf(rows)
    