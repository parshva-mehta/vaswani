import csv
from datetime import date
import os

def getTodayDate():
    return date.today()
    
def createReport():
    print("Creating report, please answer the follwing questions")
    while(True):
        row = 0
        #Inputs 
        manager = input("Task assinged by: ")
        description = input("Brief Description of task: ")
        deadline = input("Date for completion: ")
        status = input("Complete, In-Progress, Incomplete: ")
        notes = input("Additional notes: ")
        
        
        manage = []
        manage.append(manager)
        desc = []
        desc.append(description)
        due = []
        due.append(deadline)
        stat = []
        stat.append(status)
        note = []
        note.append(notes)
        
        data = []

        data.append(manage[row])
        data.append(desc[row])
        data.append(due[row])
        data.append(stat[row])
        data.append(note[row])
        # write the data

        writer.writerow(data)

 
        
        prompt_continue = input("Task created, would you like to add another task? \n 1 = Yes, 2 = No ")
        row = row + 1
        if(prompt_continue == '2'):
            break;
            file.close()

def saveReport():
   old_name = r"C:\Users\parsh\Documents\Visual Studio Code\Python\Status_report.csv"
   new_name = fr"C:\Users\parsh\\Documents\Visual Studio Code\Python\\Status_report{getTodayDate()}\.csv"
   os.rename(old_name, new_name)

def exportReport():  #WORK IN PROGRESS
    import email, smtplib, ssl
    from email import encoders
    from email.mime.base import MIMEBase
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText

    subject = "An email with attachment from Python"
    body = "This is an email with attachment sent from Python"
    sender_email = "parshvam7@gmail.com"
    receiver_email = "parshvam7@gmail.com"
    password = input("Type your password and press enter:")

    # Create a multipart message and set headers
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    message["Bcc"] = receiver_email  # Recommended for mass emails

    # Add body to email
    message.attach(MIMEText(body, "plain"))

    filename = "Status_report.csv"  # In same directory as script

    # Open PDF file in binary mode
    with open(filename, "rb") as attachment:
        # Add file as application/octet-stream
        # Email client can usually download this automatically as attachment
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    # Encode file in ASCII characters to send by email    
    encoders.encode_base64(part)

    # Add header as key/value pair to attachment part
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {filename}",
    )

    # Add attachment to message and convert message to string
    message.attach(part)
    text = message.as_string()

    # Log in to server using secure context and send email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, text)


choice = input("Please select an option: \n 1: Create a report \n 2: Edit a report \n 3: Export a report \n")
if (choice == '1'):
    columms = ['Assigned by', 'Description', 'Deadline', 'Status', 'Notes']
    with open('Status_report.csv', 'w', encoding= 'utf-8', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(columms)
        createReport()
#saveReport()
