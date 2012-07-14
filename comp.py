import win32com.client
from datetime import datetime, date
from pprint import pprint
import csv
import sys
import os.path
import subprocess
from prettytable import PrettyTable
import collections

autoit = win32com.client.Dispatch("AutoItX3.dll")
autoit.Activate(wintotal_title)

n = 0
listing = [] #records
inputFile = sys.argv[1]
with open(inputFile, 'r') as infile:
    reader = csv.DictReader(infile, delimiter=',', quotechar='"')
    for row in reader:
    
        property = Property()
        
        property.ListingDate = datetime.strptime(row['Listing_Date '], "%m/%d/%Y %H:%M:%S %p").date() if row['Listing_Date '] else None
        property.ListingPrice = float(row['Listing_Price'].replace(",","") or 0)
        property.SellingDate = datetime.strptime(row['Selling_Date'], "%m/%d/%Y %H:%M:%S %p").date() if row['Selling_Date'] else None
        property.SellingPrice = float(row['Selling_Price'].replace(",","") or 0)
        property.Ownership = row['Ownership']
        property.Age = int(scrubAge(row['Age']))
        property.Status = row['Status']
        property.DOM = int(row['DOM'] or 0)
   
        n+=1
        listing.append(property)
        
        
for i, p in enumerate(listing):
    
    #Address
    autoit.Send(p.Address + "{ENTER}")
    autoit.Send(p.City + "{ENTER}")
    autoit.Send(p.State + "{ENTER}")
    autoit.Send(p.Zip + "{ENTER}")
    
    #Proximity to Subject
    
    #