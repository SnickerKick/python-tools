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
# The usaddress package is used to break the address into individual commponents,
# and the pysusps package to verify address
#

# Import the needed libraries

import pandas as pd
import datetime
import string
import re
import usaddress

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
                       'Zip Code',
                       'Unknown']

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

# Iterate through each entry, parse and clean


for index, row in inputFile.iterrows() :
    addressString = str(row['Address 1']) + ' ' + str(row['Address 2']) + ' '\
                     + str(row['Address 3']) + ' ' + str(row['Address 4']) +  \
                     ' ' + str(row['Address 5']) + ' ' + str(row['Address 6'])

# initialize loop accumulators
    
    address = ''
    unit = ''
    city = ''
    state = ''
    zip = ''
    unknown = ''

# break address into components

    parsedAddress = usaddress.parse(addressString)
    
    for tuple in parsedAddress :
        if tuple[0] == 'nan' :
            continue
        
        if tuple[1] == 'AddressNumber' :
            address += tuple[0] + ' '
        elif tuple[1] == 'StreetName' :
            address += tuple[0] + ' '
        elif tuple[1] == 'StreetNamePreDirectional' :
            address = tuple[0] + ' ' + address
        elif tuple[1] == 'StreetNamePostType' :
            address += tuple[0] + ' '
        elif tuple[1] == 'OccupancyType' :
            unit += tuple[0] + ' '
        elif tuple[1] == 'OccupancyIdentifier' :
            unit += tuple[0] + ' '
        elif tuple[1] == 'PlaceName' :
            city += tuple[0] + ' '
        elif tuple[1] == 'StateName' :
            state += tuple[0] + ' '
        elif tuple[1] == 'ZipCode' :
            zip += tuple[0] + ' '
        else :
            unknown += tuple[1] + ':' + tuple[0] + ' '

# clean up trailing spaces
            
    address = address.strip()
    unit = unit.strip()
    city = city.strip()
    state = state.strip()
    zip = zip.strip()

# Clean up phone number
    
    allow = string.digits
    phone = re.sub('[^%s]' % allow, '', row['Phone'])
    phone = phone[:3] + '-' + phone[3:6] + '-' + phone[6:]
        
# Build the new row    
    
    vendorNewRow = {'First Name' : row['First Name'].title(),
                      'Last Name' : row['Last Name'].title(),
                      'Primary Number' : phone,
                      'Email' : row['Email'].lower(),
                      'Address' : (address + ' ' + unit).strip(),
                      'City' : city,
                      'State' : state,
                      'Zip Code' : zip,
                      'Unknown' : unknown}
                      
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












