#create text with data from spreadsheet filling the company name, state, and name
#compose an email, type the appropriate header, body
#send all drafts
#send only if date is present
#send by week??

import pandas as pd

def convert(s):
    str1 = ""
    return (str1.join(s))


df = pd.read_excel("C:/Users/parsh/Downloads/mockdata.xlsx") #change based on machine program is run on

#create a list of all store numbers under a certain week
StoreList = df["STORE #"].tolist()

#create a list of emails (WORK IN PROGRESS)
emailList = df["EMAIL ADDRESS "].tolist()

#create list of delivery dates
ETAList = df["DELIVERY DATE"].tolist()

#create list of store names
StoreNames = df["STORE NAME"].tolist()
StateList = df["STATE"].tolist()

subjectList = []
bodyList = []
nameList = []

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
    



#depends on week
min = 2
max = 16


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
