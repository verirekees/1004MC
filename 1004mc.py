from datetime import datetime, date
from pprint import pprint
import csv

listing = [] #records

class Property:
    pass

p3date = date.today()
if p3date.month > 3:
    p3date = p3date.replace(p3date.year, p3date.month-3, p3date.day)
else:
    p3date = p3date.replace(p3date.year-1, p3date.month+9, p3date.day)

#p3date.replace(p3date.year if p3date.month > 3 else p3date.year-1, p3date.month-3 if p3date.month > 3 else p3date.month+9, p3date.day)

p6date = date.today()
if p6date.month > 6:
    p6date = p6date.replace(p6date.year, p6date.month-6, p6date.day)
else:
    p6date = p6date.replace(p6date.year-1, p6date.month+6, p6date.day)

#p6date.replace(p6date.year if p6date.month > 6 else p6date.year-1, p6date.month-6 if p6date.month > 6 else p6date.month+6, p6date.day)

p12date = date.today()
p12date = p12date.replace(p12date.year-1, p12date.month, p12date.day)

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
    if x.ListingDate < p3date and (isActive(x) or x.SellingDate > p3date):
        return True
    else:
        return False
    
def isActivePrior7To12Months(x):
    if x.ListingDate < p6date and (isActive(x) or x.SellingDate > p6date):
        return True
    else:
        return False 
   
n = 0 
with open('jennings.csv', 'rb') as f:
    reader = csv.reader(f, delimiter=',')
    #f = reader.next()
    for row in reader:
    
        property = Property()
        
        property.Age = int(row[0] or 0)
        property.DOM = int(row[1] or 0)
        property.ListingDate = datetime.strptime(row[2], "%m/%d/%Y %H:%M:%S %p").date() if row[2] else None
        property.ListingPrice = float(row[3].replace(",","") or 0)
        property.SellingDate = datetime.strptime(row[4], "%m/%d/%Y %H:%M:%S %p").date() if row[4] else None
        property.SellingPrice = float(row[5].replace(",","") or 0)
        property.PendingDate = datetime.strptime(row[6], "%m/%d/%Y %H:%M:%S %p").date() if row[6] else None
        property.Status = row[7]
        property.Ownership = row[8]
   
        n+=1     
        print n, row
        listing.append(property)

t1 = len(listing)

#w1 = len(filter(isPrior7To12Months, listing))
#w2 = len(filter(isPrior4To6Months, listing))
#w3 = len(filter(isPrior3Months, listing))

x1 = len(filter(isSettledPrior7To12Months, filter(isSettled, listing)))
x2 = len(filter(isSettledPrior4To6Months, filter(isSettled, listing)))
x3 = len(filter(isSettledPrior3Months, filter(isSettled, listing)))

w1 = len(filter(isActivePrior7To12Months, listing))
w2 = len(filter(isActivePrior4To6Months, listing))
w3 = len(filter(isActivePrior3Months, listing))

print t1, x1, x2, x3, w1, w2, w3
