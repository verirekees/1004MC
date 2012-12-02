import win32com.client
from datetime import datetime, date
from pprint import pprint
import csv
import sys
import os.path
import subprocess
from prettytable import PrettyTable
import collections
import re

class Property:
    pass

def scrubAge(rawAge):
    if rawAge == "":
        return -1

    age = int(rawAge)       
    if age == 0:
        return 1
    elif age > 1800:
        return date.today().year - age

    return age
    
def splitTax(rawStr):
    m = re.search("(?<=Tax: )(.*?)(?=\s)", rawStr, re.I)
    if m:
        return m.group(1)
    else:
        return rawStr

def splitMLS(rawStr):
    m = re.search("(?<=MLS: )(.*)", rawStr, re.I)
    if m:
        return m.group(1)
    else:
        return rawStr
        
listing = [] #records

n = 0
compFile = sys.argv[1]
with open(compFile, 'r') as infile:
    reader = csv.DictReader(infile, delimiter=',', quotechar='"')
    for row in reader:

        property = Property()

        property.Status = row['Status']
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
        
# n = 0        
# compTaxFile = sys.argv[2]
# with open(compTaxFile, 'r') as infile2:
    # reader = csv.DictReader(infile2, delimiter=',', quotechar='"')
    # for row in reader:

        # property = listing[n]
        
        # property.MLSPhoto = row['MLS Photo Indicator']
        # property.MLSListing = row['MLS Listing Indicator']
        # property.Foreclosure = row['Foreclosure Indicator']
        # property.DistressedSale = row['Distressed Sale Indicator']
        # property.Address = row['Address']
        #property.ZipCode2 = row['ZIP Code']
        # property.Subdivision = row['Subdivision']
        # property.Owner = row['Owner Name']
        # property.RecordingDate = datetime.strptime(row['Recording Date'], "%m/%d/%Y").date() if row['Recording Date'] else None
        # property.SaleDate = datetime.strptime(splitTax(row['Sale Date (TAX|MLS)']), "%m/%d/%Y").date() if splitTax(row['Sale Date (TAX|MLS)']) else None
        #property.MLSSaleDate = datetime.strptime(splitMLS(row['Sale Date (TAX|MLS)']), "%m/%d/%Y").date() if splitMLS(row['Sale Date (TAX|MLS)']) else None
        # property.TotalAssessedValue = float(row['Total Assessed Value'][1:].replace(",","") or 0)
        # property.Tax = float(row['Tax Amount'][1:].replace(",","") or 0)
        # property.LandUse = splitTax(row['Land Use - Universal  (TAX|MLS)'])
        #property.MLSLandUse = splitMLS(row['Land Use - Universal  (TAX|MLS)'])
        # property.Beds = int(splitTax(row['Beds (TAX|MLS)']) or 0)
        #property.MLSBeds = int(splitMLS(row['Beds (TAX|MLS)']) or 0)
        # property.Baths = int(splitTax(row['Total Baths (TAX|MLS)']) or 0)
        #property.MLSBaths = int(splitMLS(row['Total Baths (TAX|MLS)']) or 0)        
        # property.LotSize = float(row['Lot Sq Ft'].replace(",","") or 0)
        # property.BuildingSize = float(splitTax(row['Building Sq Ft (TAX|MLS)']).replace(",","") or 0)
        #property.MLSBuildingSize = float(splitMLS(row['Building Sq Ft (TAX|MLS)']).replace(",","") or 0)
        # property.Story = float(splitTax(row['Stories (TAX|MLS)']) or 0)
        #property.MLSStory = float(splitMLS(row['Stories (TAX|MLS)']) or 0)
        # property.YearBuilt = int(splitTax(row['Year Built (TAX|MLS)']) or 0)
        # property.TaxID = row['Tax ID']
    
listing = sorted(listing, key=lambda property: property.MLNumber)

#for x in listing:
#    print(str(x.StreetNumber) + "\t" + x.StreetName + "\t" + x.Status)
    
propertyName = ""
m = re.search("^(.*) Comps Data", os.path.basename(sys.argv[1]), re.I)
if m:
    propertyName = m.group(1)
else:
    print("property name is empty\n")
    sys.exit(1)

autoit = win32com.client.Dispatch("AutoItX3.Control")

if autoit.WinExists("WinTOTAL - " + propertyName) == 0:
    print(propertyName + " wintotal does not exists\n")
    sys.exit(1)
    
autoit.WinActivate("WinTOTAL - " + propertyName)

for i, p in enumerate(listing):
    #if i != 5:
    #   continue

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

    #Verification Source
    autoit.Send("MLS Data Bank; City/Cnty Records" + "{ENTER}")

    #Sales or Financing Concessions (TODO)
    if p.Ownership == "Private":
        autoit.Send("ArmLth" + "{ENTER}")
    elif p.Ownership == "Bank" or p.Ownership == "Government":
        autoit.Send("REO" + "{ENTER}")
    elif p.Ownership == "Owner by Contract":
        autoit.Send("NonArm" + "{ENTER}")
    else:
        if p.Status == "Active":
            autoit.Send("Listing" + "{ENTER}")
        else:
            autoit.Send("{ENTER}")
    autoit.Send("{ENTER}") #adjustment column
    autoit.Send(p.Financing + "{ENTER}") #Financing Type
    autoit.Send("{ENTER}") #Other description:
    autoit.Send(str(0) + "{ENTER}") #Concession amount (TODO)
    autoit.Send("{ENTER}") #adjustment column

    #Date of Sale/Time
    if p.Status == "Active":
        autoit.Send("Active{ENTER}") #Status of Comp
        autoit.Send("{ENTER}{ENTER}") #Contract Date Boolean
        autoit.Send("{ENTER}") #Contract Date
        autoit.Send("{ENTER}") #Settlement Date
        autoit.Send("{ENTER}") #Withdrawal Date
        autoit.Send("{ENTER}") #Expiration Date
    else:
        autoit.Send("Settled sale{ENTER}") #Status of Comp
        autoit.Send("{SPACE}") #Contract Date Boolean
        autoit.Send(p.PendingDate.strftime("%m/%y")+"{ENTER}") #Contract Date
        autoit.Send(p.SellingDate.strftime("%m/%y")+"{ENTER}") #Settlement Date
        autoit.Send("{ENTER}") #Withdrawal Date
        autoit.Send("{ENTER}") #Expiration Date
    autoit.Send("{ENTER}") #adjustment column

    #Location
    autoit.Send("Neutral" + "{ENTER}") #Overall
    autoit.Send("Residential" + "{ENTER}") #Factor 1
    autoit.Send("{ENTER}") #Desc 1
    autoit.Send("{ENTER}") #Factor 2
    autoit.Send("{ENTER}") #Desc 2
    autoit.Send("{ENTER}") #adjustment column

    #Leasehold/Fee Simple
    autoit.Send("Fee Simple" + "{ENTER}")
    autoit.Send("{ENTER}") #adjustment column

    #Site
    autoit.Send("{ENTER}")
    autoit.Send("{ENTER}") #adjustment column

    #View Overall
    autoit.Send("Neutral" + "{ENTER}") #Overall
    autoit.Send("Residential" + "{ENTER}") #Factor 1
    autoit.Send("{ENTER}") #Desc 1
    autoit.Send("{ENTER}") #Factor 2
    autoit.Send("{ENTER}") #Desc 2
    autoit.Send("{ENTER}") #adjustment column

    #Design (Style)
    if re.search("sty|story", p.Style, re.I):
        autoit.Send("Traditional" + "{ENTER}")
    else:
        autoit.Send(p.Style + "{ENTER}")
    autoit.Send("{ENTER}") #adjustment column

    #Quality of Construction
    autoit.Send("=0")
    autoit.Sleep(200)
    autoit.Send("{ENTER}")
    autoit.Send("{ENTER}") #adjustment column

    #Actual Age
    autoit.Send(str(p.Age) + "{ENTER}")
    autoit.Send("{ENTER}") #adjustment column

    #Condition
    autoit.Send("=0")
    autoit.Sleep(200)
    autoit.Send("{ENTER}")
    autoit.Send("{ENTER}") #adjustment column

    #Above Grade
    autoit.Send("{ENTER}") #adjustment column

    #Room Count
    autoit.Send(str(p.Room)+ "{ENTER}") #Total
    autoit.Send(str(p.MainUpperBedrooms) + "{ENTER}") #Bedrooms
    fullbath = p.MainFullBaths + p.UpperFullBaths
    halfbath = p.MainHalfBaths + p.UpperHalfBaths
    autoit.Send(str(fullbath) + '.' + str(halfbath) + "{ENTER}") #Baths
    autoit.Send("{ENTER}") #adjustment column

    #Gross Living Area
    autoit.Send(str(p.SqFt) + "{ENTER}")
    autoit.Send("{ENTER}") #adjustment column

    #Basement & Finished Rooms Below Grade
    basementSqFt = 0
    if re.search("1.5", p.Style, re.I):
        basementSqFt = int(p.SqFt * 0.75)
    elif re.search("2", p.Style, re.I):
        basementSqFt = int(p.SqFt * 0.50)
    else:
        basementSqFt = p.SqFt
    autoit.Send(str(basementSqFt) + "{ENTER}")
    if re.search("full", p.BasementDesc, re.I):
        autoit.Send(str(basementSqFt) + "{ENTER}") #Finished Sq Ft
    elif re.search("partial", p.BasementDesc, re.I):
        autoit.Send(str(int(0.5 * basementSqFt)) + "{ENTER}") #Finished Sq Ft
    elif re.search("crawl", p.BasementDesc, re.I):
        autoit.Send("0" + "{ENTER}") #Finished Sq Ft
    else:
        autoit.Send("0" + "{ENTER}") #Finished Sq Ft
    m = re.search('Walk.*(?=[\s|,])', p.BasementDesc, re.I)
    if m:
        autoit.Send(str(m.group(0)) + "{ENTER}") #Basement Exit
    else:
        autoit.Send("Interior-only" + "{ENTER}") #Basement Exit
    autoit.Send("{ENTER}") #adjustment column
    autoit.Send("0" + "{ENTER}") #Rec Room
    autoit.Send(str(p.LowerBedrooms) + "{ENTER}") #Bedrooms
    autoit.Send(str(p.LowerFullBaths) + '.' + str(p.LowerHalfBaths) + "{ENTER}") #Bathrooms
    autoit.Send("0" + "{ENTER}") #Other Rooms
    autoit.Send("{ENTER}") #adjustment column

    #Functional Utility
    autoit.Send("Average" + "{ENTER}")
    autoit.Send("{ENTER}") #adjustment column

    #Heating/Cooling
    if re.search("Gas", p.HeatingSource, re.I) and re.search("Forced Air", p.Heating, re.I):
       if re.search("Central-Electric", p.Cooling, re.I):
           autoit.Send("GFWA/Central" + "{ENTER}")
       else:
           autoit.Send("GFWA/Cooling" + "{ENTER}")
    else:
       autoit.Send("Other" + "{ENTER}")
    autoit.Send("{ENTER}") #adjustment column

    #Energy Efficient Items
    autoit.Send("Storm Sash" + "{ENTER}")
    autoit.Send("{ENTER}") #adjustment column

    #Garage/Carport
    parkingStr = ""
    if p.GarageNum > 0 and p.CarportNum == 0:
        parkingStr = str(p.GarageNum) + " Gar/";
    elif p.GarageNum == 0 and p.CarportNum > 0:
        parkingStr = str(p.CarportNum) + " Car/";
    elif p.GarageNum > 0 and p.CarportNum > 0:
        parkingStr = str(p.GarageNum) + " Gar/" + str(p.CarportNum) + " Crp/";
    else:
        parkingStr = "None"
    if re.search("attached", p.ParkingDesc, re.I):
        parkingStr += "Attached"
    elif re.search("detached", p.ParkingDesc, re.I):
        parkingStr += "Dettached"
    elif re.search("tuck", p.ParkingDesc, re.I):
        parkingStr += "BuiltIn"
    else:
        parkingStr += ""
    autoit.Send(parkingStr + "{ENTER}")
    autoit.Send("{ENTER}") #adjustment column

    #Porch/Patio/Deck
    outsideStr = ""
    if re.search("patio", p.MISC, re.I):
        outsideStr += "Patio"
    if re.search("porch", p.MISC, re.I):
        if outsideStr == "":
            outsideStr += "Porch"
        else:
            outsideStr += "/Porch"
    if re.search("deck", p.MISC, re.I):
        if outsideStr == "":
            outsideStr += "Deck"
        else:
            outsideStr += "/Deck"
    if outsideStr == "":
        outsideStr = "None"
    autoit.Send(outsideStr + "{ENTER}")
    autoit.Send("{ENTER}") #adjustment column

    #DOM/CDOM
    domStr = str(p.DOM) if p.DOM > 0 else "n/a"
    domStr += "/" + str(p.CDOM)
    autoit.Send(domStr + "{ENTER}")
    autoit.Send("{ENTER}") #adjustment column

    #Kitchen
    if re.search("Updated", p.KitchenDesc, re.I):
        autoit.Send("Updated" + "{ENTER}")
    elif re.search("Custom Cabinetry", p.KitchenDesc, re.I):
        autoit.Send("Updated" + "{ENTER}")
    else:
        autoit.Send("Normal" + "{ENTER}")
    autoit.Send("{ENTER}") #adjustment column

    #Amenities
    amenStr = ""
    m = re.search("Fencing", p.LotDesc, re.I)
    if m and p.FireplaceNum > 0:
        amenStr = "FNC/" + str(p.FireplaceNum) + " FPL"
    elif m and p.FireplaceNum == 0:
        amenStr = "Fencing"
    elif not m and p.FireplaceNum > 0:
        amenStr = str(p.FireplaceNum) + " Fireplace"
    if amenStr != "":
        autoit.Send(amenStr + "{ENTER}")
    else:
        autoit.Send("{ENTER}")
    autoit.Send("{ENTER}") #adjustment column

    #Date of Prior S/T
    autoit.Send("{ENTER}")

    #Price of Prior S/T
    autoit.Send("{ENTER}")

    #Data Source(s)
    autoit.Send("MLS & County Records"+"{ENTER}")

    #Effective Date
    autoit.Send(date.today().strftime("%m/%d/%Y") + "{ENTER}")