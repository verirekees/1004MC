def median(s):
    i = len(s)
    if not i%2:
        return (s[int(i/2)-1]+s[int(i/2)])/2.0
    return s[int((i-1)/2)]

def roundUp(x, n=-3):
    if n < 0:
        return round(float(x)/(10 ** -n) + 0.1) * (10 ** -n)
    else:
        return round(float(x), n)
    
def pred(l):
    roundedList = map(roundUp, l);
    freqDict = collections.Counter(roundedList)
    maxFreq = max(freqDict.values())    
    maxFreqList = [v for v, f in freqDict.items() if f == maxFreq]
    return median(maxFreqList)
    
print("the predominate price is " + str(pred(price)) + "\n");
