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

def median(s):
    i = len(s)
    if i <= 0:
        return 0
    elif not i%2:
        return (s[int(i/2)-1]+s[int(i/2)])/2.0
    else:
        return s[int((i-1)/2)]

def isSettled(x):
    return x.Status == "Sold"

def isActive(x):
    return x.Status == "Active"
    
def isListedPrior3Months(x):
    if x.ListingDate >= p3date:
        return True
    else:
        return False
    
def isListedPrior4To6Months(x):
    if x.ListingDate >= p6date and x.ListingDate < p3date:
        return True
    else:
        return False
    
def isListedPrior7To12Months(x):
    if x.ListingDate >= p12date and x.ListingDate < p6date:
        return True
    else:
        return False 

def isSettledPrior3Months(x):
    if x.SellingDate >= p3date:
        return True
    else:
        return False
    
def isSettledPrior4To6Months(x):
    if x.SellingDate >= p6date and x.SellingDate < p3date:
        return True
    else:
        return False
    
def isSettledPrior7To12Months(x):
    if x.SellingDate >= p12date and x.SellingDate < p6date:
        return True
    else:
        return False 

def isActivePrior3Months(x):
    if isActive(x):
        return True
    else:
        return False
    
def isActivePrior4To6Months(x):
    if x.ListingDate <= p3date and (isActive(x) or x.SellingDate >= p3date):
        return True
    else:
        return False
    
def isActivePrior7To12Months(x):
    if x.ListingDate <= p6date and (isActive(x) or x.SellingDate >= p6date):
        return True
    else:
        return False 

def isForeclosure(x):
    if x.Ownership == 'Bank' or x.Ownership == 'Government':
        return True
    else:
        return False
        
def scrubAge(rawAge):
    if rawAge == "":
        return -1
    
    age = int(rawAge)
    if age == 0:
        return 1
    elif age > 1800:
        return date.today().year - age

    return age

def roundUp(x):
    return round(float(x)/(10 ** 3) + 0.1) * (10 ** 3)

def pred(l, roundUp = False):
    roundedList = map(roundUp, l) if roundUp == True else l;
    freqDict = collections.Counter(roundedList)
    maxFreq = max(freqDict.values())
    maxFreqList = [v for v, f in freqDict.items() if f == maxFreq]
    medianVal = median(sorted(roundedList))
    diffList = dict([(abs(medianVal-x), x) for x in maxFreqList])
    return diffList[min(diffList.keys())]


# --- Collect Data --- #
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



# --- Calculate Results --- #
p3date = date.today()
if p3date.month > 3:
    p3date = p3date.replace(p3date.year, p3date.month-3, p3date.day)
else:
    p3date = p3date.replace(p3date.year-1, p3date.month+9, p3date.day)

p6date = date.today()
if p6date.month > 6:
    p6date = p6date.replace(p6date.year, p6date.month-6, 28)#p6date.day)
else:
    p6date = p6date.replace(p6date.year-1, p6date.month+6, 28)#p6date.day)

p12date = date.today()
p12date = p12date.replace(p12date.year-1, p12date.month, p12date.day)

settledList = list(filter(isSettled, listing)) #Settled Listings

prior7To12MonthsSaleList = list(filter(isSettledPrior7To12Months, settledList))
prior4To6MonthsSaleList = list(filter(isSettledPrior4To6Months, settledList))
prior3MonthsSaleList = list(filter(isSettledPrior3Months, settledList))

prior7To12MonthsActiveList = list(filter(isActivePrior7To12Months, listing))
prior4To6MonthsActiveList = list(filter(isActivePrior4To6Months, listing))
prior3MonthsActiveList = list(filter(isActivePrior3Months, listing))

#Foreclosure Rate
foreclosureRate = len(list(filter(isForeclosure, listing))) / len(listing)

medianSalesPrice = [] #Median Comparable Sale Price
medianSalesPrice.append(median(sorted([x.SellingPrice for x in prior7To12MonthsSaleList])))
medianSalesPrice.append(median(sorted([x.SellingPrice for x in prior4To6MonthsSaleList])))
medianSalesPrice.append(median(sorted([x.SellingPrice for x in prior3MonthsSaleList])))

medianSalePriceTrend = [] #Median Sale Price Trend
if medianSalesPrice[0] < medianSalesPrice[2]:
    medianSalePriceTrend.append((medianSalesPrice[2] - medianSalesPrice[0])/medianSalesPrice[0])
    medianSalePriceTrend.append('inclined')
else:
    medianSalePriceTrend.append((medianSalesPrice[0] - medianSalesPrice[2])/medianSalesPrice[0])
    medianSalePriceTrend.append('declined')

medianSalesDOM = [] #Median Comparable Sales DOM
medianSalesDOM.append(median(sorted([x.DOM for x in prior7To12MonthsSaleList])))
medianSalesDOM.append(median(sorted([x.DOM for x in prior4To6MonthsSaleList])))
medianSalesDOM.append(median(sorted([x.DOM for x in prior3MonthsSaleList])))
medianSalesDOMTrend = 'increased' if medianSalesDOM[0] < medianSalesDOM[2] else 'decreased';

medianListPrice = [] #Median Comparable List Price
medianListPrice.append(median(sorted([x.ListingPrice for x in prior7To12MonthsActiveList])))
medianListPrice.append(median(sorted([x.ListingPrice for x in prior4To6MonthsActiveList])))
medianListPrice.append(median(sorted([x.ListingPrice for x in prior3MonthsActiveList])))

medianListingPriceTrend = [] #Median Listing Price Trend
if medianListPrice[0] < medianListPrice[1]:
    medianListingPriceTrend.append((medianListPrice[1] - medianListPrice[0])/medianListPrice[0])
    medianListingPriceTrend.append('inclined')
else:
    medianListingPriceTrend.append((medianListPrice[0] - medianListPrice[1])/medianListPrice[0])
    medianListingPriceTrend.append('declined')

medianListDOM = [] #Median Comparable Listings DOM
medianListDOM.append(median(sorted([x.DOM for x in prior7To12MonthsActiveList])))
medianListDOM.append(median(sorted([x.DOM for x in prior4To6MonthsActiveList])))
medianListDOM.append(median(sorted([x.DOM for x in prior3MonthsActiveList])))
medianListDOMTrend = 'increased' if medianListDOM[0] < medianListDOM[2] else 'decreased';

medianSaleOverList = [] #Median Sale Price as % of List Price
medianSaleOverList.append(median(sorted([x.SellingPrice/x.ListingPrice for x in prior7To12MonthsSaleList])))
medianSaleOverList.append(median(sorted([x.SellingPrice/x.ListingPrice for x in prior4To6MonthsSaleList])))
medianSaleOverList.append(median(sorted([x.SellingPrice/x.ListingPrice for x in prior3MonthsSaleList])))

marketTrend = ""
if medianSalePriceTrend[0] < -0.05:
    marketTrend = "declining"
elif medianSalePriceTrend[0] > 0.05:
    marketTrend = "inclining"
else:
    marketTrend = "stable"

salesTable = PrettyTable(["Median Sale & List Price, DOM, Sale/List %        ", "Prior 7-12 Months", "Prior 4-6 Months", "Current - 3 Months", "Overall Trend"])
salesTable.add_row(["Median Comparable Sale Price", medianSalesPrice[0], medianSalesPrice[1], medianSalesPrice[2], "NA"])
salesTable.add_row(["Median Comparable Sales Days on Market", medianSalesDOM[0], medianSalesDOM[1], medianSalesDOM[2], "NA"])
salesTable.add_row(["Median Comparable List Price", medianListPrice[0], medianListPrice[1], medianListPrice[2], "NA"])
salesTable.add_row(["Median Comparable Listings Days on Market", medianListDOM[0], medianListDOM[1], medianListDOM[2], "NA"])
salesTable.add_row(["Median Sale Price as % of List Price", '{0:.2%}'.format(medianSaleOverList[0]), '{0:.2%}'.format(medianSaleOverList[1]), '{0:.2%}'.format(medianSaleOverList[2]), "NA"])
salesTable.align = 'l'
    
sellingPriceStat = [min([x.SellingPrice for x in listing]), 
                    max([x.SellingPrice for x in listing]),
                    pred([x.SellingPrice for x in listing])]

ageStat = [min([x.Age for x in listing if x.Age > 0]),
           max([x.Age for x in listing if x.Age > 0]),
           pred([x.Age for x in listing if x.Age > 0])]

sellingPriceTable = PrettyTable(["{0:28}".format("Selling Price Statistics"), "{0:10}".format("Low"), "{0:10}".format("High"), "{0:10}".format("Predominate")])
sellingPriceTable.add_row(["Price", sellingPriceStat[0], sellingPriceStat[1], sellingPriceStat[2]])
sellingPriceTable.align = 'c'

ageTable = PrettyTable(["{0:28}".format("Age Statistics"), "{0:10}".format("Low"), "{0:10}".format("High"), "{0:10}".format("Predominate")])
ageTable.add_row(["Age", ageStat[0], ageStat[1], ageStat[2]])
ageTable.align = 'c'

outputFile = os.path.splitext(inputFile)[0] + '_sum.txt'
with open(outputFile, 'w') as outfile:
    outfile.write(salesTable.get_string())
    outfile.write("\n\n\n")
    outfile.write(
"Based on a research on all the sold properties within of sujbect property, \
the following market conditions were drawn.  \
The foreclosure rate is about {0:.0%} in this area over the last twelve months.  \
The median sale prices have {1} about {2:.0%} \
while the median list prices have {3} about {4:.0%} over the last twelve months.  \
The median days on market have slightly {5}.  \
The median sale price is as {6:.0%} of the list price in the last three months.  \
This is a {7} Market.\n\n\n\n".format(foreclosureRate, 
                                      medianSalePriceTrend[1],
                                      medianSalePriceTrend[0], 
                                      medianListingPriceTrend[1], 
                                      medianListingPriceTrend[0],
                                      medianSalesDOMTrend,
                                      medianSaleOverList[2], 
                                      marketTrend));
    outfile.write(sellingPriceTable.get_string())
    outfile.write('\n')
    outfile.write(ageTable.get_string())
    outfile.close()
    
subprocess.call(['C:/WINDOWS/System32/Notepad.exe', outputFile])