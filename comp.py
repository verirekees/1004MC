import win32com.client
from datetime import datetime, date
from pprint import pprint
import csv
import sys
import os.path
import subprocess
from prettytable import PrettyTable
import collections

class Property:
    pass

def scrubAge(rawAge):
    if rawAge == "":
        return -1

    age = int(rawAge)
    if age == 0:
        return 1
    elif age > 1900:
        return date.today().year - age

    return age

n = 0
listing = [] #records
inputFile = sys.argv[1]
with open(inputFile, 'r') as infile:
    reader = csv.DictReader(infile, delimiter=',', quotechar='"')
    for row in reader:

        property = Property()

        property.StreetNumber = int(row["Street_Number"] or 0)
        property.MLNumber = int(row["ML_Number"] or 0)
        property.StreetName = row["Street_Name"]
        property.City = row["City"]
        property.State = row["State"]
        property.ZipCode = int(row["Zip_Code"] or 0)
        property.MLS = row["MLS"]
        property.CDOM = int(row["CDOM"] or 0)
        property.Ownership = row["Ownership"]
        property.Financing = row["Financing"] if row["Financing"] else "Conventional";
        property.CC = int(row["CC"] or 0)
        property.ContigentDate = datetime.strptime(row["Contingent_Date"], "%m/%d/%Y %H:%M:%S %p").date() if row["Contingent_Date"] else None
        property.PendingDate = datetime.strptime(row["Pending_Date"], "%m/%d/%Y %H:%M:%S %p").date() if row["Pending_Date"] else None
        property.Style = row["Style"]
        property.Room = int(row["ROOM"] or 0)
        property.MainUpperBedrooms = int(row["BELL"] or 0)
        property.MainFullBaths = int(row["MNFL"] or 0)
        property.MainHalfBaths = int(row["MNHF"] or 0)
        property.UpperFullBaths = int(row["UPFL"] or 0)
        property.UpperHalfBaths = int(row["UPHF"] or 0)
        property.SqFt = int(row["Square_Footage"] or 0)
        property.BasementSqFt = int(row["Basement_Square_Footage"] or 0)
        property.BasementDesc = row["BSMT"]
        property.LowerBedrooms = int(row["LOWR"] or 0)
        property.LowerFullBaths = int(row["LWFL"] or 0)
        property.LowerHalfBaths = int(row["LWHF"] or 0)
        property.Heating = row["HEAT"]
        property.HeatingSource = row["HETS"]
        property.Cooling = row["COOL"]
        property.GarageNum = int(row["GRG_No"] or 0)
        property.CarportNum = int(row["CRP_No"] or 0)
        property.ParkingDesc = row["PRKD"]
        property.MISC = row["MISC"]
        property.KitchenDesc = row["KITC"]
        property.LotDesc = row["LTDS"]
        property.FireplaceNum = int(row["FRP_No"] or 0)
        property.FireplaceType = row["FPLD"]
        property.FireplaceLocation = row["FRPL "]
        property.ListingPrice = float(row['Listing_Price'].replace(",","") or 0)
        property.SellingDate = datetime.strptime(row['Selling_Date'], "%m/%d/%Y %H:%M:%S %p").date() if row['Selling_Date'] else None
        property.SellingPrice = float(row['Selling_Price'].replace(",","") or 0)
        property.Ownership = row['Ownership']
        property.Age = int(scrubAge(row['Age']))
        property.DOM = int(row['DOM'] or 0)

        n+=1
        listing.append(property)

#p = listing[0]
autoit = win32com.client.Dispatch("AutoItX3.Control")
autoit.WinActivate("WinTOTAL - 9877 Zenith Drive")

for i, p in enumerate(listing):
    #if i > 1:
    #    break

    #Address
    autoit.Send(str(p.StreetNumber) + ' ' + str(p.StreetName) + "{ENTER}")
    autoit.Send(str(p.City) + "{ENTER}")
    autoit.Send(str(p.State) + "{ENTER}")
    autoit.Send(str(p.ZipCode) + "{ENTER}")

    #Proximity to Subject
    autoit.Send("{ENTER}")

    #Sale Price (TODO need listing if selling not available)
    if p.SellingPrice > 0:
        autoit.Send(str(int(p.SellingPrice)) + "{ENTER}")
    else:
        autoit.Send(str(int(p.ListingPrice)) + "{ENTER}")
    
    #Data Sources
    autoit.Send("Maris{#}" + str(p.MLNumber) + "{ENTER}")
    autoit.Send(str(p.DOM) + "{ENTER}")

    # #Verification Source
    autoit.Send("MLS Data Bank; City/Cnty Records" + "{ENTER}")

    # #Sales or Financing Concessions (TODO)
    autoit.Send("{ENTER}")
    autoit.Send("{ENTER}") #adjustment column
    autoit.Send(p.Financing + "{ENTER}") #Financing Type
    autoit.Send("{ENTER}") #Other description:
    autoit.Send(str(0) + "{ENTER}") #Concession amount (TODO)
    autoit.Send("{ENTER}") #adjustment column

    # #Date of Sale/Time
    autoit.Send("Active{ENTER}") #Status of Comp
    autoit.Send("{ENTER}{ENTER}") #Contract Date Boolean
    autoit.Send("{ENTER}") #Contract Date
    autoit.Send("{ENTER}") #Settlement Date
    autoit.Send("{ENTER}") #Withdrawal Date
    autoit.Send("{ENTER}") #Expiration Date
    autoit.Send("{ENTER}") #adjustment column

    # #Location
    autoit.Send("Neutral" + "{ENTER}") #Overall
    autoit.Send("Residential" + "{ENTER}") #Factor 1
    autoit.Send("{ENTER}") #Desc 1
    autoit.Send("{ENTER}") #Factor 2
    autoit.Send("{ENTER}") #Desc 2
    autoit.Send("{ENTER}") #adjustment column

    # #Leasehold/Fee Simple
    autoit.Send("Fee Simple" + "{ENTER}")
    autoit.Send("{ENTER}") #adjustment column

    # #Site
    autoit.Send("{ENTER}")
    autoit.Send("{ENTER}") #adjustment column

    # #View Overall
    autoit.Send("Neutral" + "{ENTER}") #Overall
    autoit.Send("Residential" + "{ENTER}") #Factor 1
    autoit.Send("{ENTER}") #Desc 1
    autoit.Send("{ENTER}") #Factor 2
    autoit.Send("{ENTER}") #Desc 2
    autoit.Send("{ENTER}") #adjustment column

    # #Design (Style)
    autoit.Send(p.Style + "{ENTER}")
    autoit.Send("{ENTER}") #adjustment column

    # #Quality of Construction
    autoit.Send("=0")
    autoit.Sleep(200)
    autoit.Send("{ENTER}")
    autoit.Send("{ENTER}") #adjustment column

    # #Actual Age
    autoit.Send(str(p.Age) + "{ENTER}")
    autoit.Send("{ENTER}") #adjustment column

    # #Condition
    autoit.Send("=0")
    autoit.Sleep(200)
    autoit.Send("{ENTER}")
    autoit.Send("{ENTER}") #adjustment column

    # #Above Grade
    autoit.Send("{ENTER}") #adjustment column

    # #Room Count
    autoit.Send(str(p.Room)+ "{ENTER}") #Total
    autoit.Send(str(p.MainUpperBedrooms) + "{ENTER}") #Bedrooms
    fullbath = p.MainFullBaths + p.UpperFullBaths
    halfbath = p.MainHalfBaths + p.UpperHalfBaths
    autoit.Send(str(fullbath) + '.' + str(halfbath) + "{ENTER}") #Baths
    autoit.Send("{ENTER}") #adjustment column

    # #Gross Living Area
    autoit.Send(str(p.SqFt) + "{ENTER}")
    autoit.Send("{ENTER}") #adjustment column

    # #Basement & Finished Rooms Below Grade 
    autoit.Send(str(p.BasementSqFt) + "{ENTER}") #Area Sq Ft
    autoit.Send("{ENTER}") #Finished Sq Ft
    autoit.Send("{ENTER}") #Basement Exit
    autoit.Send("{ENTER}") #adjustment column
    autoit.Send("{ENTER}") #Rec Room
    autoit.Send(str(p.LowerBedrooms) + "{ENTER}") #Bedrooms
    autoit.Send(str(p.LowerFullBaths) + '.' + str(p.LowerHalfBaths) + "{ENTER}") #Bathrooms
    autoit.Send("{ENTER}") #Other Rooms
    autoit.Send("{ENTER}") #adjustment column

    # #Functional Utility
    autoit.Send("Average" + "{ENTER}")
    autoit.Send("{ENTER}") #adjustment column

    # #Heating/Cooling
    autoit.Send(p.Heating + "{ENTER}")
    autoit.Send("{ENTER}") #adjustment column

    # #Energy Efficient Items
    autoit.Send("Storm Sash" + "{ENTER}")
    autoit.Send("{ENTER}") #adjustment column

    # #Garage/Carport
    autoit.Send("{ENTER}")
    autoit.Send("{ENTER}") #adjustment column

    # #Porch/Patio/Deck
    autoit.Send("{ENTER}")
    autoit.Send("{ENTER}") #adjustment column

    # #DOM/CDOM
    domStr = str(p.DOM) if p.DOM > 0 else "n/a"
    domStr += "/" + str(p.CDOM)
    autoit.Send(domStr + "{ENTER}")
    autoit.Send("{ENTER}") #adjustment column

    # #Kitchen
    autoit.Send("Normal" + "{ENTER}")
    autoit.Send("{ENTER}") #adjustment column

    # #Amenities
    autoit.Send("{ENTER}")
    autoit.Send("{ENTER}") #adjustment column

    # #Date of Prior S/T
    autoit.Send("{ENTER}")

    # #Price of Prior S/T
    autoit.Send("{ENTER}")

    # #Data Source(s)
    autoit.Send("MLS & County Records"+"{ENTER}")

    # #Effective Date
    autoit.Send(date.today().strftime("%m/%d/%Y") + "{ENTER}")