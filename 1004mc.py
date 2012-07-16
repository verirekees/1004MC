from datetime import datetime, date
from pprint import pprint
import csv
import sys
import os.path
import subprocess
import matplotlib.pyplot as plt

#--- Set Effective Date ---#
effectiveDate = None
#effectiveDate = date(2012, 7, 10)

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



# --- Collect Data --- #
n = 0
listing = [] #records
inputFile = sys.argv[1]
with open(inputFile, 'r') as infile:
    reader = csv.DictReader(infile, delimiter=',', quotechar='\"')
    for row in reader:
    
        property = Property()
        
        property.DOM = int(row['DOM'] or 0)
        property.ListingDate = datetime.strptime(row['Listing_Date'], "%m/%d/%Y %H:%M:%S %p").date() if row['Listing_Date'] else None
        property.ListingPrice = float(row['Listing_Price'].replace(",","") or 0)
        property.SellingDate = datetime.strptime(row['Selling_Date'], "%m/%d/%Y %H:%M:%S %p").date() if row['Selling_Date'] else None
        property.SellingPrice = float(row['Selling_Price'].replace(",","") or 0)
        property.PendingDate = datetime.strptime(row['Pending_Date '], "%m/%d/%Y %H:%M:%S %p").date() if row['Pending_Date '] else None
        property.Status = row['Status']
   
        n+=1
        listing.append(property)



# --- Calculate Results --- #
p3date = effectiveDate or date.today()
if p3date.month > 3:
    p3date = p3date.replace(p3date.year, p3date.month-3, p3date.day)
else:
    p3date = p3date.replace(p3date.year-1, p3date.month+9, p3date.day)

p6date = date.today()
if p6date.month > 6:
    p6date = p6date.replace(p6date.year, p6date.month-6, p6date.day)
else:
    p6date = p6date.replace(p6date.year-1, p6date.month+6, p6date.day)

p12date = date.today()
p12date = p12date.replace(p12date.year-1, p12date.month, p12date.day)

settledList = list(filter(isSettled, listing)) #Settled Listings

prior7To12MonthsSaleList = list(filter(isSettledPrior7To12Months, settledList))
prior4To6MonthsSaleList = list(filter(isSettledPrior4To6Months, settledList))
prior3MonthsSaleList = list(filter(isSettledPrior3Months, settledList))

prior7To12MonthsActiveList = list(filter(isActivePrior7To12Months, listing))
prior4To6MonthsActiveList = list(filter(isActivePrior4To6Months, listing))
prior3MonthsActiveList = list(filter(isActivePrior3Months, listing))


numCompSales = [] #Number of Comparable Sales
numCompSales.append(len(prior7To12MonthsSaleList))
numCompSales.append(len(prior4To6MonthsSaleList))
numCompSales.append(len(prior3MonthsSaleList))

absorpRate = [] #Absorption Rate
absorpRate.append(numCompSales[0] / 6.0)
absorpRate.append(numCompSales[1] / 3.0)
absorpRate.append(numCompSales[2] / 3.0)

numCompActive = [] #Number of Comparable Active Listings
numCompActive.append(len(prior7To12MonthsActiveList))
numCompActive.append(len(prior4To6MonthsActiveList))
numCompActive.append(len(prior3MonthsActiveList))

numMonthsOfSupply = [] #Number of Months of Housing Supply
numMonthsOfSupply.append(numCompActive[0] / absorpRate[0] if absorpRate[0] != 0 else -1)
numMonthsOfSupply.append(numCompActive[1] / absorpRate[1] if absorpRate[1] != 0 else -1)
numMonthsOfSupply.append(numCompActive[2] / absorpRate[2] if absorpRate[2] != 0 else -1)

medianSalesPrice = [] #Median Comparable Sale Price
medianSalesPrice.append(median(sorted([x.SellingPrice for x in prior7To12MonthsSaleList])))
medianSalesPrice.append(median(sorted([x.SellingPrice for x in prior4To6MonthsSaleList])))
medianSalesPrice.append(median(sorted([x.SellingPrice for x in prior3MonthsSaleList])))

medianSalesDOM = [] #Median Comparable Sales DOM
medianSalesDOM.append(median(sorted([x.DOM for x in prior7To12MonthsSaleList])))
medianSalesDOM.append(median(sorted([x.DOM for x in prior4To6MonthsSaleList])))
medianSalesDOM.append(median(sorted([x.DOM for x in prior3MonthsSaleList])))

medianListPrice = [] #Median Comparable List Price
medianListPrice.append(median(sorted([x.ListingPrice for x in prior7To12MonthsActiveList])))
medianListPrice.append(median(sorted([x.ListingPrice for x in prior4To6MonthsActiveList])))
medianListPrice.append(median(sorted([x.ListingPrice for x in prior3MonthsActiveList])))

medianListDOM = [] #Median Comparable Listings DOM
medianListDOM.append(median(sorted([x.DOM for x in prior7To12MonthsActiveList])))
medianListDOM.append(median(sorted([x.DOM for x in prior4To6MonthsActiveList])))
medianListDOM.append(median(sorted([x.DOM for x in prior3MonthsActiveList])))

medianSaleOverList = [] #Median Sale Price as % of List Price
medianSaleOverList.append(median(sorted([x.SellingPrice/x.ListingPrice for x in prior7To12MonthsSaleList])))
medianSaleOverList.append(median(sorted([x.SellingPrice/x.ListingPrice for x in prior4To6MonthsSaleList])))
medianSaleOverList.append(median(sorted([x.SellingPrice/x.ListingPrice for x in prior3MonthsSaleList])))

# --- Print Results --- #
from prettytable import PrettyTable

def createInvAnlyTable():
    invAnlyTable = PrettyTable(["Inventory Analysis", "Prior 7-12 Months", "Prior 4-6 Months", "Current - 3 Months", "Overall Trend"])
    invAnlyTable.add_row(["Total # of Comparable Sales(Settled)", numCompSales[0], numCompSales[1], numCompSales[2], "NA"])
    invAnlyTable.add_row(["Absorption Rate (Total Sales/Months)", '{0:.1f}'.format(absorpRate[0]), '{0:.1f}'.format(absorpRate[1]), '{0:.1f}'.format(absorpRate[2]), "NA"])
    invAnlyTable.add_row(["Total # of Comparable Active Listings", numCompActive[0], numCompActive[1], numCompActive[2], "NA"])
    invAnlyTable.add_row(["Months of Housing Supply (Total Listings/Ab. Rate)", '{0:.1f}'.format(numMonthsOfSupply[0]), '{0:.1f}'.format(numMonthsOfSupply[1]), '{0:.1f}'.format(numMonthsOfSupply[2]), "NA"])
    invAnlyTable.align = 'l'
    return invAnlyTable
    
def createSalesTable():
    salesTable = PrettyTable(["Median Sale & List Price, DOM, Sale/List %        ", "Prior 7-12 Months", "Prior 4-6 Months", "Current - 3 Months", "Overall Trend"])
    salesTable.add_row(["Median Comparable Sale Price", medianSalesPrice[0], medianSalesPrice[1], medianSalesPrice[2], "NA"])
    salesTable.add_row(["Median Comparable Sales Days on Market", medianSalesDOM[0], medianSalesDOM[1], medianSalesDOM[2], "NA"])
    salesTable.add_row(["Median Comparable List Price", medianListPrice[0], medianListPrice[1], medianListPrice[2], "NA"])
    salesTable.add_row(["Median Comparable Listings Days on Market", medianListDOM[0], medianListDOM[1], medianListDOM[2], "NA"])
    salesTable.add_row(["Median Sale Price as % of List Price", '{0:.2%}'.format(medianSaleOverList[0]), '{0:.2%}'.format(medianSaleOverList[1]), '{0:.2%}'.format(medianSaleOverList[2]), "NA"])
    salesTable.align = 'l'
    return salesTable

invAnlyTable = createInvAnlyTable()
salesTable = createSalesTable()

outputFile = os.path.splitext(inputFile)[0] + '_table.txt'
with open(outputFile, 'w') as outfile:
    outfile.write(invAnlyTable.get_string())
    outfile.write('\n')
    outfile.write(salesTable.get_string())
    outfile.close()
    
subprocess.call(['C:/WINDOWS/System32/Notepad.exe', outputFile])

d = dict([(x.SellingDate, x.SellingPrice) for x in listing if x.Status == 'Sold'])

dateAxis = []
priceAxis = []
for date in sorted(d.keys()):
    dateAxis.append(date)
    priceAxis.append(d[date])
    
avgPriceAxis = []
for i, date in enumerate(dateAxis):
    avgPriceAxis.append(sum(priceAxis[0:i+1]) / len(priceAxis[0:i+1]))


#show()
plt.subplot(111)
plt.grid()
fig = plt.gcf()
fig.set_size_inches(20, 8)
p = plt.plot(dateAxis, priceAxis, dateAxis, avgPriceAxis)
plt.legend(p, ["Selling Price", "Trend(Average Price)"])

outputFile = os.path.splitext(inputFile)[0] + '_graph.png'
plt.savefig(outputFile, dpi=100)