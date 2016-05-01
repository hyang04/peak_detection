#! python3
# EXPLANATION: please see document on how to run this program to get set up.


#importing all the modules
import os
import re
import sys
import openpyxl
from openpyxl.cell import get_column_letter
import comtypes, comtypes.client
from ctypes import *
from comtypes.automation import *

# a function to find the next column, implemented to display data
def nextcolumn(curcol):
    if len(curcol) == 1:
        if curcol == 'Z':
            return 'AA'
        else:
            return chr(ord(curcol) + 1)
    else:
        lastLetter = curcol[-1]
        if lastLetter == 'Z':
            firstLetter = curcol[0]
            newfirstletter = chr(ord(firstLetter) + 1)
            return ''.join([newfirstletter, 'A'])
        else:
            newlastletter = chr(ord(lastLetter) + 1)
            return ''.join([curcol[0], newlastletter])
#  the following consists of the function, GetChroData. GetChroData is part of the MSFileReader package, so be sure you have MSFileReader ready to go.
def GetChroData(mass, pdStartTime, pdEndTime, filename): # compare to line 324
    # don't mess with the things below. we still aren't too sure how any of these things work.
    nChroOperator = 0
    nChroType2 = 0
    nChroType1 = 0
    bstrfilter = u'FTMS' 
    dDelay = 0.0
    bstrMassRanges1 = str(float(format((mass-0.02), '.2f'))) + "-" + str(float(format((mass+0.02), '.2f'))) #be within plus or minus .02 in ranges, appears as Number-Number
    bstrMassRanges2 = ""
    nSmoothingType = 0
    nSmoothingValue = 0
    pnArraySize = 0
    cd = VARIANT() 
    pf = VARIANT()

    fInName = filename # raw file
    xr = comtypes.client.CreateObject('MSFileReader.XRawfile')
    xr.open(fInName)
    res = xr.SetCurrentController(0,1)
    pnArraySize = c_long()
    xr.GetChroData(c_long(nChroType1), c_long(nChroOperator), c_long(nChroType2),
        bstrfilter,
        bstrMassRanges1,
        bstrMassRanges2,
        dDelay, c_double(pdStartTime), c_double(pdEndTime),
        c_long(nSmoothingType),
        c_long(nSmoothingValue),
        cd,
        pf,
        pnArraySize)
    if cd.value == None:
        print('The computer is unable to locate that raw data file currently. Please make sure it is in the right spot.')
    global peptideresults # setting as a global variable for using outside function
    peptideresults = list(cd.value)# works like a nested dictionary
    xr.close()


#the following detects the highest point of data
def DoPeakDetection(data, timey, RT): #compare to line 327
    peak_valley = [] #you insert p, v, u, or d into this list. p = peak, v = valley, u = up, d = down. this begins the process of finding peaks.
    justpeakvalleythings = {} #this is to isolate those p's and v's, including ones at the beginning and end
    left_valley = 0 #sets left valley at a certain time
    right_valley = len(data) #starts off at very end, moves to the right
    i = 0
    for i in range(len(data)):
        if i == 0: #starting place
            if data[i+1] > data[i]:
                peak_valley.insert(i, ['v1', timey[i]])
            else:
                peak_valley.insert(i, ['p1', timey[i]])
        elif i == len(data) - 1:
            if data[i-1] > data[i]:
                peak_valley.insert(i, ['v1', timey[i]])
            else:
                peak_valley.insert(i, ['p1', timey[i]])
        else: #middle areas
            a = data[i-1]
            b = data[i]
            c = data[i+1]
            if a < b and b < c:
                peak_valley.insert(i, ['u', timey[i]])
            elif a > b and b > c:
                peak_valley.insert(i, ['d', timey[i]])
            elif a < b and b > c:
                peak_valley.insert(i, ['p',timey[i]])
            elif a > b and b < c:
                peak_valley.insert(i, ['v', timey[i]])
            elif b == 0.0 and c == 0.0: # loops through a sequence of zeroes until it reaches the last one; when c is not zero
                peak_valley.insert(i, 'zero sequence')
            else:
                peak_valley.insert(i, ['v', timey[i]])
    for z in range(right_valley): #left_valley and right_valley will coalesce at RT (retention time)
        if timey[z] < RT:
            if peak_valley[z] == 'v' or peak_valley[z]=='v1':
                left_valley = timey[z]
            else:
                print('', end='')
        elif peak_valley[z] == 'v' or peak_valley[z]=='v1':
            right_valley = timey[z]
            break
        else:
            print('',end='')
    a = 0
    for x in range(len(peak_valley)): #as mentioned earlier, you are isolating peaks and valleys here
        if peak_valley[x][0] == 'p':
            justpeakvalleythings.setdefault(a, {'p': [data[x], timey[x]]})
            a = a + 1
        elif peak_valley[x][0] == 'v': 
            justpeakvalleythings.setdefault(a, {'v': [data[x], timey[x]]})
            a = a + 1
        elif peak_valley[x][0] == 'v1':
            justpeakvalleythings.setdefault(a, {'v1': [data[x], timey[x]]})
            a = a + 1
        elif peak_valley[x][0] == 'p1':
            justpeakvalleythings.setdefault(a, {'p1': [data[x], timey[x]]})
            a = a + 1
        else:
            print('', end='')
    global peakvalley
    peakvalley = peak_valley
    global peakvalleythings
    peakvalleythings = justpeakvalleythings
  
  
# here is where you use what you found earlier p's, v's, etc to find the highest peak  
def HighestPeakAndRT(peakvalley, data, timey, peaklist, RT, coordinate, importantstuff):
    highestpeak = [0]
    f = -1
    for keys in range(len(peaklist)):
        f = f + 1
        if list(peaklist[keys].keys())[0] == 'p':
            for variab in range(len(highestpeak)):
                if peaklist[keys]['p'][0] < highestpeak[variab]: # finding highest peak by comparing to previous highest
                    print('', end='')
                elif highestpeak[0] == 0:
                    highestpeak.insert(0, peaklist[keys]['p'][0])
                    break
                else:
                    highestpeak.insert(variab, peaklist[keys]['p'][0])
                    break
            if variab == len(highestpeak)-1:
                highestpeak.insert(0, peaklist[keys]['p'][0]) # so this is for you to insert the first peak as the highest, to do the above for loop
            else:
                print('',end='')
        elif list(peaklist[keys].keys())[0] == 'p1': #the same as before, but looking at the beginning or ending peaks
            for variab in range(len(highestpeak)):
                if peaklist[keys]['p1'][0] < highestpeak[variab]:
                    print('', end='')
                elif highestpeak[0] == 0:
                    highestpeak.insert(0, peaklist[keys]['p1'][0])
                    break
                else:
                    highestpeak.insert(variab, peaklist[keys]['p1'][0])
                    break
            if variab == len(highestpeak)-1:
                highestpeak.insert(0, peaklist[keys]['p1'][0])
            else:
                print('',end='')
        else:
            print('', end='')
    for ind in range(len(highestpeak)): 
        peakHeight = highestpeak[ind]
        c = 0
        leftboundary = []
        rightboundary = []
        for b in range(len(data)): # here you are seeing finding that peak's place
            if data[c] == peakHeight:
                break
            else:
                c = c + 1     
        for index in range(c, len(data)): # you are finding the boundaries of the peak
            if rightboundary != []: # see below
                break
            else:
                # setting rightboundary at the first lowest point
                if peakvalley[index][0] == 'v' and data[index] <= peakHeight * 0.005 :
                    rightboundary = [timey[index], data[index]]
                    break
                elif peakvalley[index][0] == 'v1' and data[index] <= peakHeight * 0.005:
                    rightboundary = [timey[index], data[index]]
                    break
                elif peakvalley[index][0] == 'd' and data[index] <= peakHeight * 0.005:
                    rightboundary = [timey[index], data[index]]
                    break
                elif peakvalley[index][0] == 'u' and data[index] <= peakHeight * 0.005:
                    rightboundary = [timey[index], data[index]]
                    break
                elif peakvalley[index][0] == 'p' and data[index] <= peakHeight * 0.005:
                    rightboundary = [timey[index], data[index]]
                    break
                elif peakvalley[index][0] == 'v' and data[index] >= peakHeight * 0.005: # for the case when the boundary isn't low enough, do peak separation and find a peak that's .25 the size
                    for indy in range(index, len(data)):
                        if peakvalley[indy][0] == 'p' or peakvalley[indy][0] == 'p1':
                            if peakHeight > data[indy]:
                                division = data[indy]/peakHeight
                                if division > .25:
                                    rightboundary = [timey[index], data[index]]
                                    break
                                else:
                                    print('', end='')
                            else:
                                division = peakHeight/data[indy]
                                if division > .25:
                                    rightboundary = [timey[index], data[index]]
                                    break
                                else:
                                    print('', end='')
                        else: 
                            print('',end='')
                elif peakvalley[index][0]=='v1' and data[index] >= peakHeight * 0.005:
                    for indy in range(index, len(data)):
                        if peakvalley[indy][0] == 'p' or peakvalley[indy][0] == 'p1':
                            if peakHeight > data[indy]:
                                division = data[indy]/peakHeight
                                if division > .25:
                                    rightboundary = [timey[index], data[index]]
                                    break
                                else:
                                    print('', end='')
                            else:
                                division = peakHeight/data[indy]
                                if division > .25:
                                    rightboundary = [timey[index], data[index]]
                                    break
                                else:
                                    print('', end='')
                        else: 
                            print('',end='')
                else:
                    print('', end='')   
            if index == len(data)-1 and rightboundary==[]:
                rightboundary = [timey[index], data[index]] 
            else:
                print('', end='')           
        for inde in reversed(range(0, c-1)): #same as above, but with leftboundary
            if leftboundary != []:
                break
            else:
                if peakvalley[inde][0] == 'v' and data[inde] <= peakHeight * 0.005 :
                    leftboundary = [timey[inde], data[inde]] 
                    break
                elif peakvalley[inde][0] == 'v1' and data[inde] <= peakHeight * 0.005:
                    leftboundary = [timey[inde], data[inde]]
                    break
                elif peakvalley[inde][0] == 'd' and data[inde] <= peakHeight * 0.005:
                    leftboundary = [timey[inde], data[inde]]
                    break
                elif peakvalley[inde][0] == 'u' and data[inde] <= peakHeight * 0.005:
                    leftboundary = [timey[inde], data[inde]]
                    break
                elif peakvalley[inde][0] == 'p' and data[inde] <= peakHeight * 0.005:
                    leftboundary = [timey[inde], data[inde]]
                    break
                elif peakvalley[inde][0] == 'v' and data[inde] >= peakHeight * 0.005:
                    for indy in reversed(range(0, inde)):
                        if peakvalley[indy][0] == 'p' or peakvalley[indy][0] == 'p1':
                            if peakHeight > data[indy]:
                                division = data[indy]/peakHeight
                                if division > .25:
                                    leftboundary = [timey[inde], data[inde]]
                                    break
                                else:
                                    print('', end='')
                            else:
                                division = peakHeight/data[indy]
                                if division > .25:
                                    leftboundary = [timey[inde], data[inde]]
                                    break
                                else:
                                    print('', end='')
                        else: 
                            print('',end='')
                elif peakvalley[inde][0]=='v1' and data[inde] >= peakHeight * 0.005:
                    for indy in reversed(range(0, inde)):
                        if peakvalley[indy][0] == 'p' or peakvalley[indy][0] == 'p1':
                            if peakHeight > data[indy]:
                                division = data[indy]/peakHeight
                                if division > .25:
                                    leftboundary = [timey[inde], data[inde]]
                                    break
                                else:
                                    print('', end='')
                            else:
                                division = peakHeight/data[indy]
                                if division > .25:
                                    leftboundary = [timey[inde], data[inde]]
                                    break
                                else:
                                    print('', end='')
                        else: 
                            print('',end='')
                else:
                    print('', end='') 
            if inde == 0 and leftboundary == [0]:
                leftboundary = [timey[inde], data[inde]]
            else:
                print('', end='')
        #if neither work, set the whole thing as a boundary
        if leftboundary == []:
            leftboundary = [timey[0], data[0]]
        else:
            print('', end='')
        if rightboundary == []:
            rightboundary = [timey[len(timey)-1], data[len(data)-1]]
        else:
            print('', end='')
        if RT < rightboundary[0] and RT > leftboundary[0]: #setting the coordinate in the excel file
            importantstuff.setdefault(coordinate.coordinate, [rightboundary, leftboundary, peakHeight] )
            sheet[PeakHeight + str(coordinate.row)] = peakHeight
            sheet[LeftBoundary + str(coordinate.row)] = str(leftboundary)
            sheet[RightBoundary + str(coordinate.row)] = str(rightboundary)
            if sheet[TotalPeakHeight + str(coordinate.row)].value == None:
                sheet[TotalPeakHeight + str(coordinate.row)] = peakHeight
            else:
                peakvalue = sheet[TotalPeakHeight + str(coordinate.row)].value
                sheet[TotalPeakHeight + str(coordinate.row)] = peakvalue + peakHeight
            if trialz == 2:
                sheet[TwoPeakHeight + str(coordinate.row)] = peakHeight
            elif trialz == 3:
                sheet[ThreePeakHeight + str(coordinate.row)] = peakHeight
            elif trialz == 4:
                sheet[FourPeakHeight + str(coordinate.row)] = peakHeight
            elif trialz == 5:
                sheet[FivePeakHeight + str(coordinate.row)] = peakHeight
            else:
                sheet[SixPeakHeight + str(coordinate.row)] = peakHeight
            return [leftboundary[0], rightboundary[0], peakHeight]
        else:
            print('', end='')
            
def findHighestPeakAndRTForZValues(peakvalley, data, timey, peaklist, RT, coordinate, importantstuff):
    highestpeak = [0]
    f = -1
    for keys in range(len(peaklist)):
        f = f + 1
        if list(peaklist[keys].keys())[0] == 'p':
            for variab in range(len(highestpeak)):
                if peaklist[keys]['p'][0] < highestpeak[variab]: # finding highest peak by comparing to previous highest
                    print('', end='')
                elif highestpeak[0] == 0:
                    highestpeak.insert(0, peaklist[keys]['p'][0])
                    break
                else:
                    highestpeak.insert(variab, peaklist[keys]['p'][0])
                    break
            if variab == len(highestpeak)-1:
                highestpeak.insert(0, peaklist[keys]['p'][0]) # so this is for you to insert the first peak as the highest, to do the above for loop
            else:
                print('',end='')
        elif list(peaklist[keys].keys())[0] == 'p1': #the same as before, but looking at the beginning or ending peaks
            for variab in range(len(highestpeak)):
                if peaklist[keys]['p1'][0] < highestpeak[variab]:
                    print('', end='')
                elif highestpeak[0] == 0:
                    highestpeak.insert(0, peaklist[keys]['p1'][0])
                    break
                else:
                    highestpeak.insert(variab, peaklist[keys]['p1'][0])
                    break
            if variab == len(highestpeak)-1:
                highestpeak.insert(0, peaklist[keys]['p1'][0])
            else:
                print('',end='')
        else:
            print('', end='')
    for ind in range(len(highestpeak)): 
        peakHeight = highestpeak[ind]
        c = 0
        leftboundary = []
        rightboundary = []
        for b in range(len(data)): # here you are seeing finding that peak's place
            if data[c] == peakHeight:
                break
            else:
                c = c + 1     
        for index in range(c, len(data)): # you are finding the boundaries of the peak
            if rightboundary != []: # see below
                break
            else:
                # setting rightboundary at the first lowest point
                if peakvalley[index][0] == 'v' and data[index] <= peakHeight * 0.005 :
                    rightboundary = [timey[index], data[index]]
                    break
                elif peakvalley[index][0] == 'v1' and data[index] <= peakHeight * 0.005:
                    rightboundary = [timey[index], data[index]]
                    break
                elif peakvalley[index][0] == 'd' and data[index] <= peakHeight * 0.005:
                    rightboundary = [timey[index], data[index]]
                    break
                elif peakvalley[index][0] == 'u' and data[index] <= peakHeight * 0.005:
                    rightboundary = [timey[index], data[index]]
                    break
                elif peakvalley[index][0] == 'p' and data[index] <= peakHeight * 0.005:
                    rightboundary = [timey[index], data[index]]
                    break
                elif peakvalley[index][0] == 'v' and data[index] >= peakHeight * 0.005: # for the case when the boundary isn't low enough, do peak separation and find a peak that's .25 the size
                    for indy in range(index, len(data)):
                        if peakvalley[indy][0] == 'p' or peakvalley[indy][0] == 'p1':
                            if peakHeight > data[indy]:
                                division = data[indy]/peakHeight
                                if division > .25:
                                    rightboundary = [timey[index], data[index]]
                                    break
                                else:
                                    print('', end='')
                            else:
                                if data[indy] == 0:
                                    print("divide by zero")
                                division = peakHeight/data[indy]
                                if division > .25:
                                    rightboundary = [timey[index], data[index]]
                                    break
                                else:
                                    print('', end='')
                        else: 
                            print('',end='')
                elif peakvalley[index][0]=='v1' and data[index] >= peakHeight * 0.005:
                    for indy in range(index, len(data)):
                        if peakvalley[indy][0] == 'p' or peakvalley[indy][0] == 'p1':
                            if peakHeight > data[indy]:
                                division = data[indy]/peakHeight
                                if division > .25:
                                    rightboundary = [timey[index], data[index]]
                                    break
                                else:
                                    print('', end='')
                            else:
                                division = peakHeight/data[indy]
                                if division > .25:
                                    rightboundary = [timey[index], data[index]]
                                    break
                                else:
                                    print('', end='')
                        else: 
                            print('',end='')
                else:
                    print('', end='')   
            if index == len(data)-1 and rightboundary==[]:
                rightboundary = [timey[index], data[index]] 
            else:
                print('', end='')           
        for inde in reversed(range(0, c-1)): #same as above, but with leftboundary
            if leftboundary != []:
                break
            else:
                if peakvalley[inde][0] == 'v' and data[inde] <= peakHeight * 0.005 :
                    leftboundary = [timey[inde], data[inde]] 
                    break
                elif peakvalley[inde][0] == 'v1' and data[inde] <= peakHeight * 0.005:
                    leftboundary = [timey[inde], data[inde]]
                    break
                elif peakvalley[inde][0] == 'd' and data[inde] <= peakHeight * 0.005:
                    leftboundary = [timey[inde], data[inde]]
                    break
                elif peakvalley[inde][0] == 'u' and data[inde] <= peakHeight * 0.005:
                    leftboundary = [timey[inde], data[inde]]
                    break
                elif peakvalley[inde][0] == 'p' and data[inde] <= peakHeight * 0.005:
                    leftboundary = [timey[inde], data[inde]]
                    break
                elif peakvalley[inde][0] == 'v' and data[inde] >= peakHeight * 0.005:
                    for indy in reversed(range(0, inde)):
                        if peakvalley[indy][0] == 'p' or peakvalley[indy][0] == 'p1':
                            if peakHeight > data[indy]:
                                division = data[indy]/peakHeight
                                if division > .25:
                                    leftboundary = [timey[inde], data[inde]]
                                    break
                                else:
                                    print('', end='')
                            else:
                                division = peakHeight/data[indy]
                                if division > .25:
                                    leftboundary = [timey[inde], data[inde]]
                                    break
                                else:
                                    print('', end='')
                        else: 
                            print('',end='')
                elif peakvalley[inde][0]=='v1' and data[inde] >= peakHeight * 0.005:
                    for indy in reversed(range(0, inde)):
                        if peakvalley[indy][0] == 'p' or peakvalley[indy][0] == 'p1':
                            if peakHeight > data[indy]:
                                division = data[indy]/peakHeight
                                if division > .25:
                                    leftboundary = [timey[inde], data[inde]]
                                    break
                                else:
                                    print('', end='')
                            else:
                                division = peakHeight/data[indy]
                                if division > .25:
                                    leftboundary = [timey[inde], data[inde]]
                                    break
                                else:
                                    print('', end='')
                        else: 
                            print('',end='')
                else:
                    print('', end='') 
            if inde == 0 and leftboundary == [0]:
                leftboundary = [timey[inde], data[inde]]
            else:
                print('', end='')
        #if neither work, set the whole thing as a boundary
        if leftboundary == []:
            leftboundary = [timey[0], data[0]]
        else:
            print('', end='')
        if rightboundary == []:
            rightboundary = [timey[len(timey)-1], data[len(data)-1]]
        else:
            print('', end='')
        if RT < rightboundary[0] and RT > leftboundary[0]: #setting the coordinate in the excel file
            importantstuff.setdefault(coordinate.coordinate, [rightboundary, leftboundary, peakHeight] )
            if sheet[TotalPeakHeight + str(coordinate.row)].value == None:
                sheet[TotalPeakHeight + str(coordinate.row)] = peakHeight
            else: 
                peakvalues = sheet[TotalPeakHeight + str(coordinate.row)].value
                sheet[TotalPeakHeight + str(coordinate.row)] = peakvalues + peakHeight
            if trialz == 2:
                sheet[TwoPeakHeight + str(coordinate.row)] = peakHeight
            elif trialz == 3:
                sheet[ThreePeakHeight + str(coordinate.row)] = peakHeight
            elif trialz == 4:
                sheet[FourPeakHeight + str(coordinate.row)] = peakHeight
            elif trialz == 5:
                sheet[FivePeakHeight + str(coordinate.row)] = peakHeight
            else:
                sheet[SixPeakHeight + str(coordinate.row)] = peakHeight
            break
        else:
            print('', end='')
    return peakHeight
            
def processVariousZ(moverzxl, rowOfCellObjects, trialz, realzvalue, rtxl, runnamexl, re, leftBoundary, rightBoundary, peakvalley, peakvalleythings):
    importantstuff = {}
    peptide = sheet[moverzxl + str(rowOfCellObjects)].value
    NEEDED = [peptide]
    coordinate = sheet[moverzxl + str(rowOfCellObjects)]
    peptidemass = float(sheet[moverzxl + str(rowOfCellObjects)].value)
    experimentz = float(trialz)
    neutralmass = (peptidemass * float(realzvalue)) - (1.007825 * float(realzvalue))
    trialpeptidemass = (neutralmass + 1.007825 * experimentz) / experimentz
    # print("Row ", rowOfCellObjects, ", peptide mass = ", peptidemass, ", neutral mass = ", neutralmass, ", trial z = ", trialz, ", trial peptide mass", trialpeptidemass)
    peptideRT = sheet[rtxl + str(rowOfCellObjects)].value
    filename = sheet[runnamexl + str(rowOfCellObjects)].value
    runnameregex = re.compile(r'[^_]*')
    runname = runnameregex.search(filename)
    realrunname = runname.group()
    # if rowOfCellObjects == 8 and trialz == 6:
    #    print("start debugging")
    peptideresult = GetChroData(trialpeptidemass, float(leftBoundary), float(rightBoundary), realrunname) #compare to line 16
    timey = list(peptideresults[0])
    data = list(peptideresults[1])
    peakDetection = DoPeakDetection(data, timey, peptideRT) # compare to line 54
    
    highestpeaknrt = findHighestPeakAndRTForZValues(peakvalley, data, timey, peakvalleythings, peptideRT, coordinate, importantstuff)
    return highestpeaknrt
    # peakarea = PeakArea.PeakArea(importantstuff, data, timey)

#prepping the excel file and all its variables
os.chdir(sys.argv[1]) # changing directory to your cwd
wb = openpyxl.load_workbook(sys.argv[2]) # opening excel file for XL peptide 
sheet = wb.get_sheet_by_name(sys.argv[3]) # opens excel sheet

#titles
titles = sheet.rows[2]
for cell in titles:
    if cell.value == 'm/z':
        moverzxl = cell.column
    elif cell.value == 'z':
        zxl = cell.column
    elif cell.value == 'Fraction':
        runnamexl = cell.column
    elif cell.value == 'RT':
        rtxl = cell.column

        
lastxl = titles[-1].column
nextcol = nextcolumn(lastxl)
sheet[nextcol + '3'] = 'Peak Height with given z value'
PeakHeight = sheet[nextcol + '3'].column
nextcol = nextcolumn(nextcol)
sheet[nextcol + '3'] = 'Left Boundary'
LeftBoundary = sheet[nextcol + '3'].column
nextcol = nextcolumn(nextcol)
sheet[nextcol + '3'] = 'Right Boundary'
RightBoundary = sheet[nextcol + '3'].column
nextcol = nextcolumn(nextcol)
sheet[nextcol + '3'] = 'Peak Height with z values from 2 to 6'
TotalPeakHeight = sheet[nextcol + '3'].column
nextcol = nextcolumn(nextcol)
sheet[nextcol + '3'] = 'Peak Height with z value of 2'
TwoPeakHeight = sheet[nextcol + '3'].column
nextcol = nextcolumn(nextcol)
sheet[nextcol + '3'] = 'Peak Height with z value of 3'
ThreePeakHeight = sheet[nextcol + '3'].column
nextcol = nextcolumn(nextcol)
sheet[nextcol + '3'] = 'Peak Height with z value of 4'
FourPeakHeight = sheet[nextcol + '3'].column
nextcol = nextcolumn(nextcol)
sheet[nextcol + '3'] = 'Peak Height with z value of 5'
FivePeakHeight = sheet[nextcol + '3'].column
nextcol = nextcolumn(nextcol)
sheet[nextcol + '3'] = 'Peak Height with z value of 6'
SixPeakHeight = sheet[nextcol + '3'].column
nextcol = nextcolumn(nextcol)
sheet[nextcol + '3'] = 'Percentage Calculation'
highestRow = sheet.get_highest_row()

#below you are setting the values to find peakheight and boundaries

for rowOfCellObjects in range(4, highestRow+1):
    realzvalue = int(sheet[zxl + str(rowOfCellObjects)].value)
    trialz = realzvalue

    importantstuff = {}
    peptide = sheet[moverzxl + str(rowOfCellObjects)].value
    NEEDED = [peptide]
    coordinate = sheet[moverzxl + str(rowOfCellObjects)]
    peptidemass = float(sheet[moverzxl + str(rowOfCellObjects)].value)
    peptideRT = sheet[rtxl + str(rowOfCellObjects)].value
    filename = sheet[runnamexl + str(rowOfCellObjects)].value
    runnameregex = re.compile(r'[^_]*')
    runname = runnameregex.search(filename)
    realrunname = runname.group()
    peptideresult = GetChroData(peptidemass, peptideRT - 1.0, peptideRT + 1.0, realrunname) #compare to line 16
    timey = list(peptideresults[0])
    data = list(peptideresults[1])
    peakDetection = DoPeakDetection(data, timey, peptideRT) # compare to line 54
    highestpeaknrt = HighestPeakAndRT(peakvalley, data, timey, peakvalleythings, peptideRT, coordinate, importantstuff)
    leftBoundary = highestpeaknrt[0]
    rightBoundary = highestpeaknrt[1]
    realzPeakHeight = highestpeaknrt[2]

    for trialz in range(2, 7):
        if trialz == realzvalue:
            continue

        # processVariousZ(moverzxl, rowOfCellObjects, trialz, realzvalue, rtxl, runnamexl, re, leftBoundary, rightBoundary, peakvalley, peakvalleythings)
        importantstuff = {}
        peptide = sheet[moverzxl + str(rowOfCellObjects)].value
        NEEDED = [peptide]
        coordinate = sheet[moverzxl + str(rowOfCellObjects)]
        peptidemass = float(sheet[moverzxl + str(rowOfCellObjects)].value)
        experimentz = float(trialz)
        neutralmass = (peptidemass * float(realzvalue)) - (1.007825 * float(realzvalue))
        trialpeptidemass = (neutralmass + 1.007825 * experimentz) / experimentz
        # print("Row ", rowOfCellObjects, ", peptide mass = ", peptidemass, ", neutral mass = ", neutralmass, ", trial z = ", trialz, ", trial peptide mass", trialpeptidemass)
        peptideRT = sheet[rtxl + str(rowOfCellObjects)].value
        filename = sheet[runnamexl + str(rowOfCellObjects)].value
        runnameregex = re.compile(r'[^_]*')
        runname = runnameregex.search(filename)
        realrunname = runname.group()
        # if rowOfCellObjects == 8 and trialz == 6:
        #    print("start debugging")
        peptideresult = GetChroData(trialpeptidemass, float(leftBoundary), float(rightBoundary), realrunname) #compare to line 16
        timey = list(peptideresults[0])
        data = list(peptideresults[1])
        peakDetection = DoPeakDetection(data, timey, peptideRT) # compare to line 54
    
        highestpeaknrt = findHighestPeakAndRTForZValues(peakvalley, data, timey, peakvalleythings, peptideRT, coordinate, importantstuff)

    # trialz = realzvalue - 1
    # while trialz >= 2:    
    #     trialzPeakHeight = processVariousZ(moverzxl, rowOfCellObjects, trialz, realzvalue, rtxl, runnamexl, re, leftBoundary, rightBoundary, peakvalley, peakvalleythings)
    #     if trialzPeakHeight < realzPeakHeight * 0.85:
    #         break
        
    #     trialz = trialz - 1            
        # peakarea = PeakArea.PeakArea(importantstuff, data, timey)
    # trialz = realzvalue + 1
    # while trialz < 7:    
    #    trialzPeakHeight = processVariousZ(moverzxl, rowOfCellObjects, trialz, realzvalue, rtxl, runnamexl, re, leftBoundary, rightBoundary, peakvalley, peakvalleythings)
    #    if trialzPeakHeight < realzPeakHeight * 0.85:
    #        break
    #    trialz = trialz + 1            
        # peakarea = PeakArea.PeakArea(importantstuff, data, timey)

print('Now check your excel file')
wb.save(sys.argv[4]) # the new excel sheet you wanna save it as



        
