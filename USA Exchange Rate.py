#!/usr/bin/env python
# coding: utf-8

# In[1]:


# import libraries
import time
import os
import pandas as pd
import xlwings as xw
from pytz import timezone
from bs4 import BeautifulSoup
from datetime import datetime,date
from urllib.request import urlopen

#open yahoo finance url for bs4
url = 'https://sg.finance.yahoo.com/currencies'
html = urlopen(url)
#parse it to the bs4 html container
soup = BeautifulSoup(html, 'html.parser')

#
names=[]
prices=[]
for i in range(40, 404, 14):
    # find all lines in html
    for listing in soup.find_all('tr', attrs={'data-reactid':i}):
        # find all tabs in lines and append in lists
        for name in listing.find_all('td', attrs={'data-reactid':i+3}):
            names.append(name.text)
        for price in listing.find_all('td', attrs={'data-reactid':i+4}):
            prices.append(price.text)

# create dataframe to store the lists
currency=pd.DataFrame({'Date':None,"Names": names, "Prices": prices})
# get value for USA-INDIA Exchange
india_usa = currency[currency['Names']=='USD/INR']
# set current date
curr_Date = date.today().strftime("%d-%m-%Y")
india_usa['Date'] = curr_Date

# create list to append to excel file
for index, row in india_usa.iterrows():
    append_list = [row.Date,row.Names,row.Prices]
    

# if excel file exist append else create and append
file_name = "USA_INDIA_EXCHANGE RATE.xlsx"
if os.path.exists(file_name):
    #load excel file
    workbook = xw.Book(file_name)
    #   get sheet index 0 
    worksheet = workbook.sheets['Sheet1']
    rows = worksheet.range('A' + str(worksheet.cells.last_cell.row)).end('up').row
    #append data
    worksheet.range("A"+str(rows+1)).value = append_list
    # save and close()
    workbook.save()
    workbook.close()
    
else:
    workbook = xw.Book()
    worksheet = workbook.sheets['Sheet1']
    worksheet.range("A1").value = ['Date','Names','Rate']
    rows = worksheet.range('A' + str(worksheet.cells.last_cell.row)).end('up').row
    worksheet.range("A"+str(rows+1)).value = append_list
    workbook.save(file_name)
    workbook.close()


# In[ ]:




