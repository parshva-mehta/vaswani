#create text with data from spreadsheet filling the company name, state, and name
#compose an email, type the appropriate header, body
#send all drafts
#send only if date is present
#send by week??

import pandas as pd

df = pd.read_excel("C:/Users/Parshva/Documents/VSCode/mockdata.xlsx") #change based on machine program is run on
print (df)