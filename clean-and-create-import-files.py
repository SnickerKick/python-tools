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
#   - QBO Vendor Import.cvs
#   - QBO Incentives Import.cvs
#


# set up the Google Geocoding API key.  This needs to be done before importing 
# the geocoder library

import os
os.environ["GOOGLE_API_KEY"] = "AIzaSyBwqdWMfytQAuwLzG5MXmgZ9oxLbKYzTxY"

# Import the needed libraries

import geocoder as geo
import pandas as pd

from tkinter import *
from tkinter.filedialog import askopenfilename, asksaveasfilename, askdirectory

# Open the input file

filename = askopenfilename(title = 'Select Input File')

inputFile = pd.read_csv(filename)

col_headings = ['First Name',
                'Last Name',
                'Street Address',
                'Unit',
                'City',
                'State',
                'Zip',
                'Email',
                'Phone']

outputData = pd.DataFrame(columns=col_headings)

# Iterate through each entry and have Google clean up the address

for index, row in inputFile.iterrows() :
    address_string = row['Street Address'] + ' ' + str(row['City']) + ' ' + str(row['State']) + ' ' + str(row['Zip'])
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
        
# Build the new row    
    
    new_row = {'First Name' : row['First Name'].title(),
               'Last Name' : row['Last Name'].title(),
               'Street Address' : houseNumber + ' ' + street,
               'Unit' : row['Unit'],
               'City' : city,
               'State' : state,
               'Zip' : zip,
               'Email' : row['Email'].lower(),
               'Phone' : row['Phone'],
               'Incentive' : row['Incentive']}
    
    outputData = outputData.append(new_row, ignore_index=True)


# Write out the results

outputFile = asksaveasfilename(title = 'Save As..')

writer = pd.ExcelWriter(outputFile)
outputData.to_excel(writer, 'Cleansed Addresses')
writer.save()