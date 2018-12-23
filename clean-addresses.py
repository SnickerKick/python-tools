#
# clean-address.py
#
# Python script which takes a csv file with a list of raw mailing addresses
# uses the Google geocoder API to clean then break into the following fields:
#
#   - House number
#   - Street
#   - Apt / Unit / etc
#   - City
#   - State
#   - Zip Code
#
# and then write the results to a csv file
#

# set up the Google Geocoding API key.  This needs to be done before importing 
# the geocoder library

import os
os.environ["GOOGLE_API_KEY"] = "AIzaSyBwqdWMfytQAuwLzG5MXmgZ9oxLbKYzTxY"

# Import the needed libraries

import geocoder as geo
import pandas as pd

from tkinter import filedialog
from tkinter import *

# Open the input file

filename = filedialog.askopenfilename(title = 'Select Input File')

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

outputFile = pd.DataFrame(columns=col_headings)

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
               'Email' : row['email'].lower(),
               'Phone' : row['Phone']}
    
    outputFile = outputFile.append(new_row, ignore_index=True)


# Write out the results

outputFile = filedialog.asksaveasfile(title = 'Save As..')

writer = pd.ExcelWriter(outputFile)
outputFile.to_excel(writer, 'Cleansed Addresses')
writer.save()