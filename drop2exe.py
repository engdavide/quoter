import sys
import os
import pandas as pd
import numpy as np
import win32com.client as win32


def QQtoExcel(name):
    #Format and split input from csv into item list and pan list
    output = QQtoPD(name)
    itemsRaw = output[0]
    pansRaw = output[1]
    
    #Import conversion table, set index to PANEL
    file='trimConv.csv'
    conv=pd.read_csv(file)
    conv.set_index('PANEL', inplace=True)
    
    #Initialize conversion df: add SKU column, and fill out Typ column with panel type
    itemsConv = itemsRaw
    numItems = len(itemsConv['Item'])
    panType = itemsConv['Item'].values[0]
    itemsConv['Type'] = panType
    itemsConv['SKU'] = ""
    
    #Loop to turn all Panel Names in the Item column to 'PAN' Also sums up LF of panels and lumps into one column
    numPanCols = 0 
    linFt = 0
    for i in range(numItems):
    
        if itemsConv.loc[i, 'Type'] == itemsConv.loc[i, 'Item']:
            itemsConv.loc[i, 'Item'] = 'PAN'
            linFt = linFt + itemsConv.loc[i, 'Qty']
            if numPanCols != 0:
                itemsConv.drop(itemsConv.index[i], axis=0, inplace=True)
            numPanCols += 1
    
    #Update Qty of panels to total linFt iterated above
    itemsConv.loc[0,'Qty'] = linFt
    
        
    # Reset index and re-calc numItems counter
    itemsConv.reset_index(drop=True, inplace=True)
    numItems = len(itemsConv['Item'])
    
    #Loop to convert SKUs
    for i in range(numItems): 
        #Error handling because we've dorked with the index, and the numItems counter is off
        try:
            type = itemsConv.loc[i,'Type']
            item = itemsConv.loc[i,'Item']
            itemsConv.loc[i, 'SKU'] = conv.loc[type,item]
        except KeyError:
            pass
    
    
    #--Add other items, Z flashing, screws, butyl, etc--
    
    # Check if standing seam (sSeam = 1)
    if itemsConv.loc[0,'Item'] == 'GL' or 'GS' or 'GL':
            sSeam = 1
    else:
        sSeam = 0
    
    # For sSeam, add Z flashing to match # of HC, RC, EF, SW, and EW
    # For sSeam, add PS to match # of PV, TF
    # For sSeam, add pancakes for 
    
    # Reset index and re-calc numItems counter
    itemsConv.reset_index(drop=True, inplace=True)
    numItems = len(itemsConv['Item'])
    
    
    #--Loop to match SKUs...this may not need an actual loop
    
    # Reset index and re-calc numItems counter
    itemsConv.reset_index(drop=True, inplace=True)
    numItems = len(itemsConv['Item'])
    
    #Loop to convert SKUs
    for i in range(numItems): 
        #Error handling because we've dorked with the index, and the numItems counter is off
        try:
            type = itemsConv.loc[i,'Type']
            item = itemsConv.loc[i,'Item']
            itemsConv.loc[i, 'SKU'] = conv.loc[type,item]
        except KeyError:
            pass
    
    #Loop to count ZF needs and PS needs
    numZF = 0
    numPS = 0
    
    if sSeam == 1:
        for i in range(numItems):
            temp = itemsConv.loc[i,'Item']
            if temp == 'HC' or temp =='RC' or temp== 'EF' or temp == 'SW' or temp == 'EW':
                numZF = numZF + itemsConv.loc[i, 'Qty']
            if temp == 'PV' or temp == 'TF':
                numPS = numPS + itemsConv.loc[i, 'Qty']
        itemsConv.loc[numItems + 1] = [numZF,'ZF', panType, '']
        itemsConv.loc[numItems + 2] = [numPS,'PS', panType, '']
        itemsConv.loc[numItems + 3] = [round((numPS + numZF)/5)+1,'BUTYL', panType, 'BUTYL']
    
        #Loop to convert SKUs
        #Reset index and get numItems counter
    itemsConv.reset_index(drop=True, inplace=True)
    numItems = len(itemsConv['Item'])
    for i in range(numItems): 
        #Error handling because we've dorked with the index, and the numItems counter is off
        try:
            type = itemsConv.loc[i,'Type']
            item = itemsConv.loc[i,'Item']
            itemsConv.loc[i, 'SKU'] = conv.loc[type,item]
        except KeyError:
            pass
        
    print(itemsConv)
    # --Write output from ItemsConv to Excel--
    # This is the win32 version

    
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Item('MultiQuoter2018v0')
    excel.DisplayAlerts = False
    ws = wb.Worksheets("MULTIQUOTER")
    ws.Range('B14').Value = "test"
    
    #Reset index and get numItems counter
    itemsConv.reset_index(drop=True, inplace=True)
    numItems = len(itemsConv['Item'])
    
    #Write data from itemsConv to the open excel sheet
    for i in range(numItems):
        tempQty = itemsConv['Qty'].values[i]
        tempCode = itemsConv['SKU'].values[i]
        if i == 0:
            ws.Range('B14').Value = tempQty
            ws.Range('C14').Value = tempCode
        if i > 0:
            tempRow= 16 + i
            tempQtyLoc = 'B' + str(tempRow)
            tempCodeLoc = 'C' + str(tempRow)
            ws.Range(tempQtyLoc).Value = tempQty
            ws.Range(tempCodeLoc).Value=tempCode






def QQtoPD(file):
    # import data. Rename columns. Count # of rows
    df = pd.read_csv(file)
    df.columns = ['Item', 'Qty', 'Notes', 'Misc']
    dfCnt = df.shape
    cntRows = dfCnt[0]
    
    
    #Convert to NP array to make some things easier...
    dfNp = df.values
    #print(type(dfNp))
    #print(type(dfNp[0,0]))
    
    #note locations of sections...
    locSection = []
    for i in range(cntRows):
        tempStr = dfNp[i,0]
        if type(tempStr) != float:
            if 'SECTION' in tempStr:
                locSection.append(i)
    
            
    
    # of sections
    cntSection = len(locSection)
    
    #add cntRows as the "end" of the last section 
    #cntSection does not include this "end" value
    locSection.append(cntRows + 1)
    
    #For loop to reformat CSV from df to dfOut
    
    listOut = []
    panList = []
    
    #Add panel data
    for i in range(cntRows):
        if i in locSection:
            temp = {'Qty': int(df.iloc[i+1,1]), 'Item':df.iloc[i+1,0]}
            listOut.append(temp)
            
    # pull trims, skipping first row of labels and ending 1 row prior to first "SECTION" header
    for i in range(cntRows):
        if i > 0 and i < locSection[0] - 1:
            temp = {'Qty': int(df.iloc[i,1]), 'Item': df.iloc[i,0]}
            listOut.append(temp)
    
    
    
    #Generate pan list, based on loc
    for i in range(cntSection):
        for j in range(locSection[i]+1, locSection[i+1]-2):
            str = df.iloc[j,2]
            try:
                locQty = str.index('@')
                tempQty = int(str[:locQty-1])
            except ValueError:
                tempqty = 0
            
            try:
                locFeet = str.index('\'')
                tempFt = int(str[locQty+2:locFeet])
            except ValueError:
                tempFt = 0
            try:
                locIn = str.index('\"')
                tempIn = int(str[locFeet+2:locIn])
            except ValueError:
                tempIn = 0
            tempLen = tempFt * 12 + tempIn
            temp = {'Qty':tempQty, 'Feet':tempFt, 'In':tempIn, 'Length':tempLen}
            panList.append(temp)
    
    #convert to df and reorder columns
    dfOut = pd.DataFrame.from_records(listOut)
    dfOut = dfOut[['Qty','Item']]
    
    dfPanList = pd.DataFrame.from_records(panList)
    dfPanList = dfPanList[['Qty', 'Feet','In', 'Length']]
    return (dfOut, dfPanList)



droppedFile = sys.argv[1]

QQtoExcel(droppedFile)
input()

