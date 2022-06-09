#Python program to capture Linkdeln Profiles of users given their Name and Company
#Made in collaboration with Hiral Patel 

import csv
import time
import webbrowser
from cgitb import html
from io import StringIO

import requests
from bs4 import BeautifulSoup
from selenium import webdriver

result = []
url = 'https://www.google.com/search'

f = open('firstname.csv')
first_name = [item for item in csv.reader(f)]
f.close()

f = open('lastname.csv')
last_name = [item for item in csv.reader(f)]
f.close()

f = open('company_name.csv')
company = [item for item in csv.reader(f)]
f.close()


column_name = ["FirstName","LastName","Company Name","LinkedIn Profile"]


with open('linkedin_profile.csv', 'w', encoding='UTF8', newline='') as f:
    writer = csv.writer(f)
    # write the header
    writer.writerow(column_name)
 
    for i in range(len(last_name)):
        search = f'{first_name[i]} {last_name[i]} {company[i]} linkedin' # turn this to first name + last name + company + location
        
        headers = {
            'Accept' : '*/*',

            'Accept-Language': 'en-US,en;q=0.5',

            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.82',
        }

        parameters = {'q': search}
        content = requests.get(url, headers = headers, params = parameters).text
        soup = BeautifulSoup(content, 'html.parser')
        search = soup.find(id = 'search')


        first_link = search.find('a')

        if (first_link is None):
            linkedin_link = 'NA'
        else:
            linkedin_link = first_link['href']

        data = []
        data.append(first_name[i])
        data.append(last_name[i])
        data.append(company[i])
        data.append(linkedin_link)
        
        # write the data
        writer.writerow(data)
        
