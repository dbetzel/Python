####################
#author: dawn.betzel
#date created:  2021-10-19
#purpose: create a monitoring script that takes as input two files, one from
#Alex and one from Pam, cleans the data and joins the two datasets together
#then create output files for 5 different logical requirements
####################
#updates
#author: dawn.betzel
#2021-10-25
#purpose:  added in multi day writing to the output files and the ability
#to start writing a new file again on the first of the month
####################
#updates
#author: dawn.betzel
#2021-10-22
#purpose:  Chris Hoover added requirement (dfMatchingPOs['SAP_ACCOUNT'] == "0000200920")
#for Gnarlywood, Merchandise, NonPostable and DirectShip
####################

#import modules
import pandas as pd
import numpy as np
import openpyxl
from datetime import date
#set options to clear a warning -- related to python 3.9
pd.options.mode.chained_assignment = None #default='warn'
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

#First Requirements
#Remove some things that Alex coded into his file
def process1():
    #1 exclude all docTypes = 'AB', AC','SA'
    df['AlexCode'] = df['Document Type'].isin(['AB','AC','SA'])
    indexNames = df[df['Document Type'].isin(['AB','AC','SA'])].index
    df.drop(indexNames, inplace=True)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)
    print('process1 completed')

#First Requirements
#Not in use
def process2():
    #2 string match run first
    dfStringMatch = df.duplicated(subset=['Abs','WBSText','Company Code','G/L Account','Profit Center'],keep=False)
    df['StringMatch1'] = dfStringMatch
    dftempStringMatch = df[df['StringMatch1'] == True]
    print('process2 completed')

#First Requirements
#Not in use
def process3():
    #string match run second
    dfStringMatch = df.duplicated(subset=['Abs','Company Code','G/L Account','Profit Center'],keep=False)
    df['StringMatch2'] = dfStringMatch
    dftempStringMatch = df[df['StringMatch2'] == True]
    print('process3 completed')

#First Requirements
#Not in use
def process4(StringMatch):
    #4 this is our attempt to help our friends at MES to know how to properly cross reference things
    #tagging of StringMatch, marking which one goes with what
    #this should also work for duplicates to stamp off match, only difference in logic is dups would be same sign
    dfTempString = df[df[StringMatch] == True]
    dfTempString = dfTempString.sort_values(by=['Abs','POtext','Reference','Year/month'],kind='mergesort')
        
    dfTempString['MatchID'] = -1

    cou = 0
    firstComp = -1
    secondComp = -1
    dfTempString.drop(['Year/month'], axis=1)
    for row in dfTempString.index:
        if cou % 2 == 0:
            firstComp = row
        else:
            secondComp = row
            if dfTempString['Amount in local currency'].iloc[cou-1] + dfTempString['Amount in local currency'].iloc[cou] == 0:
                dfTempString['MatchID'].iloc[cou] = firstComp
                dfTempString['MatchID'].iloc[cou-1] = secondComp
                dfTempString['MatchIDBool'].iloc[cou-1] = True
        cou = cou + 1

    df['MatchID']=dfTempString['MatchID']
    dfTempString.to_excel("outputStringMatch.xlsx", sheet_name='Sheet_name_1')
    print('process4 completed')

#First Requirements
#Not in use
def process5():
    #5 find duplicates 
    dfDuplicates = df.duplicated(subset=['Year/month','Entry Date','Document Type','Company Code','G/L Account',
                                     'Profit Center','Amount in local currency'],keep=False)

    df['Duplicate'] = dfDuplicates
    dfTempDups = df[df['Duplicate'] == True]
    dfTempDups.to_excel("outputDuplicates.xlsx", sheet_name='Sheet_name_1')
    print('process5 completed')

#First Requirements
#Not in use
def process6():
    #Sonopress with PO
    dfSonopress = df
    dfSonopress['Vendor Description'].replace('', np.nan, inplace=True)
    dfSonopress.dropna(subset=['Vendor Description'], inplace=True)
    dfTempSonopressWPO = dfSonopress[(dfSonopress['Vendor Description']=='SONOPRESS GMBH') & (dfSonopress['POtext'].str[:2] == 'PO')]
    dfTempSonopressWPO['SonopressWPO'] = True
    df['SonopressWPO'] = dfTempSonopressWPO['SonopressWPO']
    dfTempSonopressWPO.to_excel("outputSonopressWPO.xlsx", sheet_name='Sheet_name_1')
    print('process6 completed')

#First Requirements
#Not in use
def process7():
    #Sonopress no PO
    dfSonopress = df
    dfSonopress['Vendor Description'].replace('', np.nan, inplace=True)
    dfSonopress.dropna(subset=['Vendor Description'], inplace=True)
    dfTempSonopress = dfSonopress[(dfSonopress['Vendor Description']=='SONOPRESS GMBH') & (dfSonopress['POtext'].str[:2] != 'PO')]
    dfTempSonopress['Sonopress'] = True
    df['Sonopress'] = dfTempSonopress['Sonopress']
    dfTempSonopress.to_excel("outputSonopress.xlsx", sheet_name='Sheet_name_1')
    print('process7 completed')

#First Requirements
#Not in use
#set column ProcessNum to the correct number for each process that had been run
def processNum():
    conditions = [
        (df['AlexCode'] == True),
        (df['StringMatch1'] == True),
        (df['StringMatch2'] == True),
        (df['SonopressWPO'] == True),
        (df['Sonopress'] == True),
        (df['Duplicate'] == True),
        (df['MatchIDBool'] > 0)
        ]
    choices = [1,2,3,6,7,5,4]
    df['ProcessNum'] = np.select(conditions, choices, default = -1)
    
#Process the file, because of the data formattting of the files there is much tranforming that needs to be done
def processFile():
    #adding columns and setting defaults
    df['POtext'] = 'abc'
    df['POTextRef'] = 'ghi'
    df['WBSText']= 'def'
    df['Abs'] = 0
    df['AddInverse'] = -1
    df['MatchID'] = -1
    df['ProcessNum'] = -1
    df['Duplicate'] = False
    df['SonopressWPO'] = False
    df['Sonopress'] = False
    df['AlexCode'] = False
    df['StringMatch1'] = False
    df['StringMatch2'] = False
    df['MatchIDBool'] = False

    #dropping columns with no data
    df.drop(columns=['Cost Center','Segment','Trading Partner','Assignment'], axis=1)

    #declare variables
    colonPO = ':PO'
    POhash = 'PO#'
    po = 'PO'
    dollar = '$'
    openParen = '('
    por = 'POR'

    print('Starting to process the input file')
    totalRows1 = len(df.index)
    for row in df.index:
        poText = df['Text'].iloc[row]
        refText = str(df['Reference'].iloc[row])
        invText = ''
        textLen = len(poText)
        if dollar in poText:
            if openParen in poText:
                result = poText.index(dollar)
                poText = poText[0:result-2]
            else:
                result = poText.index(dollar)
                poText = poText[0:result-1]
        elif colonPO in poText: 
            invText = poText[0:14]
            poText = poText[-10:]
        elif POhash in poText: #take left 14 for WBSText
            invText = poText[0:14]
            poText = 'PO' + poText[-8:]
        elif po in poText:
            invText = poText[0:14]
            poText = 'PO' + poText[-8:]
        #per Sarah McFarland do not remove POR data, just remove the 'POR'
        elif refText[0:3] == 'POR':
            refText = refText[3:]
        elif refText[0:2] == 'PO':
            refText = refText[2:]
        df['POTextRef'].iloc[row] = refText
        df['POtext'].iloc[row] = poText
        df['WBSText'].iloc[row] = invText
        df['Abs'] = abs(df['Amount in local currency'])
        df['AddInverse'] = df['Amount in local currency'] * -1
  
    #overwrite WBSText with data from WBS column if not null
    dfWBS = df
    dfWBS['WBS element'].replace('', np.nan, inplace=True)
    #overwrite POtext 
    dfRefPO = df
    dfRefPO['POTextRef'].replace('', np.nan, inplace=True)
    #same as SQL Coalesce, find the first non null value
    df['WBSText'] = np.where(df['WBS element'].isnull(),df['WBSText'],df['WBS element'])
    df['POtext'] = np.where(df['POTextRef'].isnull(),df['POtext'],df['POTextRef'])
    print('processing file completed, getting ready for comparing')

#2nd Requirements
#inner join Alex (SAP)file with Pam (MES)file
def getAllMatchingPosFromPam():
    print('GetAllMatchingPos')
    #set df1 to df, so that you can leave df fully intact
    df1 = df
    #dropna would not work even though the value of nan was added using replace() and np.nan in the processFile function
    df1.drop(df1[df1['POtext'] == 'nan'].index, inplace = True)
    #removing the non-numeric data as it will kill the call to pd.to_numeric below
    df1 = df1[df1.POtext.apply(lambda x: x.isnumeric())]
    #change the POtext column to be numeric, so you can merge the two datasets on POtext and PO_NUMBER
    df1['POtext'] = pd.to_numeric(df1['POtext'])
    #pandas equivalent of a SQL inner join
    dfMatchingPOs = pd.merge(df1, dfPam, how='inner', left_on = 'POtext', right_on = 'PO_NUMBER')
    dfMatchingPOs.to_excel("outputMatchingPOs.xlsx", sheet_name='Sheet_name_1')

#2nd Requirements
#find all POs marked as Gnarlywood that are in the 200920 account, first pass
def gnarlywood():
    dfGnarlywood = dfMatchingPOs[(dfMatchingPOs['WAREHOUSE_IDENTIFIER'] == 'GNAR') &(dfMatchingPOs['CONFIG_KEY'] != 'MH') & (dfMatchingPOs['SAP_ACCOUNT'] == "0000200920")]
    indexNames = dfMatchingPOs[(dfMatchingPOs['WAREHOUSE_IDENTIFIER'] == 'GNAR') &(dfMatchingPOs['CONFIG_KEY'] != 'MH') & (dfMatchingPOs['SAP_ACCOUNT'] == "0000200920")].index
    dfMatchingPOs.drop(indexNames, inplace = True)
    print('Gnarlywood completed')
    return dfGnarlywood

#2nd Requirements
#find all POs for merchandise that are in the 200920 account, second pass
def merchandise():
    dfMerchandise = dfMatchingPOs[(dfMatchingPOs['CONFIG_KEY'] == 'MH') & (dfMatchingPOs['SAP_ACCOUNT'] == "0000200920")]
    indexNames = dfMatchingPOs[(dfMatchingPOs['CONFIG_KEY'] == 'MH') & (dfMatchingPOs['SAP_ACCOUNT'] == "0000200920")].index
    dfMatchingPOs.drop(indexNames, inplace = True)
    print('Merchandise completed')
    return dfMerchandise

#2nd Requirements
#find all POs that are nonPostable that are in the 200920 account, third pass
def nonPostable():
    dfNonPostable = dfMatchingPOs[(dfMatchingPOs['POSTABLE_FLAG'] == 'N') & (dfMatchingPOs['SAP_ACCOUNT'] == "0000200920")]
    indexNames = dfMatchingPOs[(dfMatchingPOs['POSTABLE_FLAG'] == 'N') & (dfMatchingPOs['SAP_ACCOUNT'] == "0000200920")].index
    dfMatchingPOs.drop(indexNames, inplace = True)
    return dfNonPostable

#2nd Requirements
#find all POs that are directShip that are in the 200920 account, fourth pass
def directShip():
    dfDirectShip = dfMatchingPOs[(dfMatchingPOs['WAREHOUSE_IDENTIFIER'] == 'DIR') & (dfMatchingPOs['CONFIG_KEY'] != 'MH') & (dfMatchingPOs['POSTABLE_FLAG'] == 'Y') & (dfMatchingPOs['SAP_ACCOUNT'] == "0000200920")]
    indexNames = dfMatchingPOs[(dfMatchingPOs['WAREHOUSE_IDENTIFIER'] == 'DIR') & (dfMatchingPOs['CONFIG_KEY'] != 'MH') & (dfMatchingPOs['POSTABLE_FLAG'] == 'Y') & (dfMatchingPOs['SAP_ACCOUNT'] == "0000200920")].index
    dfMatchingPOs.drop(indexNames, inplace = True)
    print('directShip completed')
    return dfDirectShip

#2nd Requirements
#at first Sarah wanted this to run on a daily basis the code is written for this
#it should check for today.day == 1 then only run that section on the first, every other day it would run and concat the file to make for a cumulative month by the time it is done
#requirements shifted and now Chris says run only one day in the middle of the month
#Chris keeps changing the day so I currently am looking for a pattern and just changing today.day to what ever the date is that I run
def outputFiles():
    if today.day == 11:
        dfGnarlywood.to_excel("outputGnarlywood.xlsx", sheet_name='Sheet_name_1')
        dfMerchandise.to_excel("outputMerchandise.xlsx", sheet_name='Sheet_name_1')
        dfNonPostable.to_excel("outputNonPostable.xlsx", sheet_name='Sheet_name_1')
        dfDirectShip.to_excel("outputDirectShip.xlsx", sheet_name='Sheet_name_1')
        #Else in requirements, output dfMatchingPOs after all other processes have run and all logic to remove rows
        dfMatchingPOs.to_excel("outputElseMatchingPOs.xlsx", sheet_name='Sheet_name_1')
    else:
        dfGnarlywoodHold = pd.read_excel(r'C:\Users\DawnBetzel\OneDrive - Warner Music Group\Projects\FY2021\MES\outputGnarlywood.xlsx')
        resultGnarlywood = pd.concat([dfGnarlywoodHold, dfGnarlywood], ignore_index=True, sort=False)
        resultGnarlywood.to_excel("outputGnarlywood.xlsx", sheet_name='Sheet_name_1')
        dfMerchandiseHold = pd.read_excel(r'C:\Users\DawnBetzel\OneDrive - Warner Music Group\Projects\FY2021\MES\outputMerchandise.xlsx')
        resultMerchandise = pd.concat([dfMerchandiseHold, dfMerchandise], ignore_index=True, sort=False)
        resultMerchandise.to_excel("outputMerchandise.xlsx", sheet_name='Sheet_name_1')
        dfNonPostableHold = pd.read_excel(r'C:\Users\DawnBetzel\OneDrive - Warner Music Group\Projects\FY2021\MES\outputNonPostable.xlsx')
        resultNonPostable = pd.concat([dfNonPostableHold, dfNonPostable], ignore_index=True, sort=False)
        resultNonPostable.to_excel("outputNonPostable.xlsx", sheet_name='Sheet_name_1')
        dfDirectShipHold = pd.read_excel(r'C:\Users\DawnBetzel\OneDrive - Warner Music Group\Projects\FY2021\MES\outputDirectShip.xlsx')
        resultDirectShip = pd.concat([dfDirectShipHold, dfNonPostable], ignore_index=True, sort=False)
        resultDirectShip.to_excel("outputDirectShip.xlsx", sheet_name='Sheet_name_1')
        dfElseMatchingPOsHold = pd.read_excel(r'C:\Users\DawnBetzel\OneDrive - Warner Music Group\Projects\FY2021\MES\outputElseMatchingPOs.xlsx')
        resultElseMatchingPOs = pd.concat([dfElseMatchingPOsHold, dfMatchingPOs], ignore_index=True, sort=False)
        resultElseMatchingPOs.to_excel("outputElseMatchingPOs.xlsx", sheet_name='Sheet_name_1')

    df.to_excel("outputAll.xlsx", sheet_name='Sheet_name_1')
    print('output Excel files completed')
    
#data sources   
df = pd.read_excel(r'C:\Users\DawnBetzel\OneDrive - Warner Music Group\Projects\FY2022\MES\10.27-11.17 ALEX SAP DATA.XLSX')
dfPam = pd.read_excel(r'C:\Users\DawnBetzel\OneDrive - Warner Music Group\Projects\FY2022\MES\10.27-11.17 PAM MES DATA.xlsx')
#will have to do this for three files and append to the bottom of df
#MES can only export one year at a time??  Why is this need to know??

today = date.today()

processFile()
process1()
#process2()
#process3()
#process4('StringMatch1')
#process4('StringMatch2')
#process5()
#process6()
#process7()
#processNum()
getAllMatchingPosFromPam()
dfMatchingPOs = dfPam
dfGnarlywood = gnarlywood()
dfMerchandise = merchandise()
dfNonPostable = nonPostable()
dfDirectShip = directShip()
outputFiles()


#fuzzy matching 25% inverse
#dfSonopressTest = df['Vendor Description'].str.contains('SONOPRESS GMBH')
#df['Sonopress'] = dfSonopressTest
#dfDistinct = df[['Document Type','Company Code','G/L Account','Profit Center','Abs']]
#dfDistinct['Abs'].drop_duplicates().sort_values()

