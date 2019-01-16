#
# clean-and create-import-files.py
#
# Python script which takes a csv file with the following columns (order is unimportant):
#   - First Name
#   - Last Name
#   - Address 1  |
#   - Address 2  |
#   - Address 3  | - The combined values of these six fields is the respondents'
#   - Address 4  | - uncleansed address.  These values can all be in one field
#   - Address 5  | - or braken across all three
#   - Address 6  |
#   - Phone
#   - Email
#   - Incentive
#
# The Google geocoder API to clean and produce two files:
#   - QBO Vendor Import.xlsx
#   - QBO Incentives Import.xlsx
#


# set up the Google Geocoding API key.  This needs to be done before importing 
# the geocoder library

import os
os.environ["GOOGLE_API_KEY"] = "AIzaSyBwqdWMfytQAuwLzG5MXmgZ9oxLbKYzTxY"

# Import the needed libraries

import geocoder as geo
import pandas as pd
import datetime
import string
import re

from tkinter.filedialog import askopenfilename, asksaveasfilename

# Open the input file

filename = askopenfilename(title = 'Select Input File')

inputFile = pd.read_csv(filename)

# Set up the two output datasets

vendorColHeadings = ['First Name',
                       'Last Name',
                       'Primary Number',
                       'Email',
                       'Address',
                       'City',
                       'State',
                       'Zip Code']

incentivesColHeadings = ['Bill Number',
                           'Vendor',
                           'Bill Date',
                           'Due Date',
                           'Terms',
                           'Memo',
                           'Expense Account',
                           'Expense Description',
                           'Expense Line Amount',
                           'Expense Customer']

vendorData = pd.DataFrame(columns=vendorColHeadings)
incentiveData = pd.DataFrame(columns=incentivesColHeadings)

# Set some variable before iterating

billNumber = 1001
billDate = datetime.datetime.now().strftime("%m/%d/%Y")

# Iterate through each entry and have Google clean up the address


for index, row in inputFile.iterrows() :
    address_string = str(row['Address 1']) + ' ' + str(row['Address 2']) + ' '\
                     + str(row['Address 3']) + ' ' + str(row['Address 4']) +  \
                     ' ' + str(row['Address 5']) + ' ' + str(row['Address 6'])
    g = geo.google(address_string)

# Handle missing data
    
    try:
        houseNumber = g.osm['addr:housenumber']
    except:
        houseNumber = '<missing houseNumber>'
        
    try:
        street = g.osm['addr:street']
    except:
        street = '<missing street>'
        
    try:
        city = g.osm['addr:city']
    except:
        city = '<missing city>'
        
    try:
        state = g.osm['addr:state']
    except:
        state = '<state>'
        
    try:
        zip = g.osm['addr:postal']
    except:
        zip = '<missing postal'

# Clean up phone number
    
    allow = string.digits
    phone = re.sub('[^%s]' % allow, '', row['Phone'])
    phone = phone[:3] + '-' + phone[3:6] + '-' + phone[6:]
        
# Build the new row    
    
    vendorNewRow = {'First Name' : row['First Name'].title(),
                      'Last Name' : row['Last Name'].title(),
                      'Primary Number' : phone,
                      'Email' : row['Email'].lower(),
                      'Address' : houseNumber + ' ' + street,
                      'City' : city,
                      'State' : state,
                      'Zip Code' : zip}
                      
    incentivesNewRow = {'Bill Number' : billNumber,
                          'Vendor' : row['First Name'].title() + ' ' + row['Last Name'].title(),
                          'Bill Date' : billDate,
                          'Due Date' : '',
                          'Terms' : 'Net 30',
                          'Memo' : '',
                          'Expense Account' : '',
                          'Expense Description' : 'Research Participants',
                          'Expense Line Amount' : row['Incentive'],
                          'Expense Customer' : ''}

    billNumber += 1
    
    vendorData = vendorData.append(vendorNewRow, ignore_index=True)
    incentiveData = incentiveData.append(incentivesNewRow, ignore_index=True)


# Write out the results

outputFile = asksaveasfilename(title = 'Vendors Save As..')

writer = pd.ExcelWriter(outputFile)
vendorData.to_excel(writer, 'Vendor Import')
writer.save()

outputFile = asksaveasfilename(title = 'Incentives Save As..')

writer = pd.ExcelWriter(outputFile)
incentiveData.to_excel(writer, 'Cleansed Addresses')
writer.save()












