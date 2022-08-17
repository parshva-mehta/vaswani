
#send all drafts
#send only if date is present
#send by week??


from email.message import EmailMessage
from optparse import Values
import smtplib
import ssl
import pandas as pd
import PySimpleGUI as sg

email_sender = 'parshvam7@gmail.com'
email_reciever = 'parshvam7@gmail.com'
email_password = 'iqnt kqiz xvvw kpdl' # [mfoh brao qunx vfrg] for parshva@vaswaniinc.com

def createWindow(subject, body):
    layout = [  [sg.Text("Accept email?")],
                [sg.Text("Subject: " + subject)],
                [sg.Text("Body: " + body)],
                [sg.Button("Accept")], 
                [sg.Button("Deny")]
             ]
    window = sg.Window("Email Preview", layout , margins=(250, 250))
    while True:
        event, Values = window.read()
        if event == "Accept" or event == "Deny" or event == sg.WIN_CLOSED:
            if event == "Accept":
                em = EmailMessage()
                em['From'] = email_sender
                em['To'] = email_reciever
                em['Subject'] = subject
                em.set_content(body)

                context = ssl.create_default_context()

                with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
                    smtp.login(email_sender, email_password)
                    smtp.sendmail(email_sender, email_reciever, em.as_string())
            break
    window.close()



def convert(s):
    str1 = ""
    return (str1.join(s))

def main():
    df = pd.read_excel("C:/Users/Parshva/Documents/VSCode/mockdata.xlsx") #change based on machine program is run on

    #create a list of all store numbers under a certain week
    StoreList = df["STORE #"].tolist()

    #create a list of emails (WORK IN PROGRESS)
    emailList = df["EMAIL ADDRESS"].tolist()

    #create list of delivery dates
    ETAList = df["DELIVERY DATE"].tolist()

    #create list of store names
    StoreNames = df["STORE NAME"].tolist()
    StateList = df["STATE"].tolist()

    subjectList = []
    bodyList = []
    nameList = []

    #depends on week
    min = 2
    max = 16

 
    for i in emailList:
        splitEmail = i.split(".")
        name = splitEmail[0]
        if(name != "N/A - Store was called"):
            nameChar = ([*name])
            nameChar[0] = nameChar[0].upper()
            nameList.append(convert(nameChar))
        else:
            name = " "
            nameList.append(name)
        





    for i in range(min, max-1):
        subject = f"{StoreList[i]}, Macy's at {StoreNames[i]}, {StateList[i]} - TOYS R US Fixture Installation Vendor Urgent"
        subjectList.append(subject)

    for i in range(min, max-1):
        body = f"""Hello {nameList[i]},
                I am reaching out regarding the TOYS R US fixture installation that will be taking place at your store on {ETAList[i]} 
                Could you kindly assist in providing the following information:
                    1) Who is the main point of contact that can be reached on the morning of delivery?
                    2) Please provide a name and direct phone number?
                    3) Please also provide the Dock Captain's name and number?
                    4) Is the dock available at 7am and are there any dock restrictions?
                    5) Are there specific instructions on where our team should be entering the building from?
                    6) Has the space inside the store where the fixtures are going to be installed been cleared out?
                    7) Is it possible to provide a picture of this area?
                    8) Are there any special restrictions or callouts that the delivery/install team need to be made aware of?
                
                Please make sure a People Leader is available during the installation to help guide the team on where the fixtures should be placed on the sales floor. \n
                Our installation team will be there at 7am to begin the offloading process.
                Please help ensure the space is cleared, the dock is available to deliver fixtures through and that the freight elevator is functioning. \n
                Please respond to all on-copy. Looking forward to hearing from you. 
            Thank you
                """
        bodyList.append(body)


    #Access email
    
    
    subject1 = subject
    body1 = body

    #em = EmailMessage()
    #em['From'] = email_sender
    #em['To'] = email_reciever
    #em['Subject'] = subject1
    #em.set_content(body1)



    createWindow(subject1, body1)
    #with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
    #    smtp.login(email_sender, email_password)
    #   smtp.sendmail(email_sender, email_reciever, em.as_string())



if(__name__ == "__main__"):
    main()

