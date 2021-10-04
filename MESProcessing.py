import pandas as pd
import numpy as np
pd.options.mode.chained_assignment = None #default='warn'
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

def process1():
    #1 exclude all docTypes = 'AC','SA'
    df['AlexCode'] = df['Document Type'].isin(['AC','SA'])
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)
    print('process1 completed')

def process2():
    #2 string match run first
    dfStringMatch = df.duplicated(subset=['Abs','WBSText','Company Code','G/L Account','Profit Center'],keep=False)
    df['StringMatch1'] = dfStringMatch
    dftempStringMatch = df[df['StringMatch1'] == True]
    print('process2 completed')

def process3():
    #string match run second
    dfStringMatch = df.duplicated(subset=['Abs','Company Code','G/L Account','Profit Center'],keep=False)
    df['StringMatch2'] = dfStringMatch
    #print('dfStringMatch',dfStringMatch)
    dftempStringMatch = df[df['StringMatch2'] == True]
    print('process3 completed')

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


def process5():
    #5 find duplicates 
    dfDuplicates = df.duplicated(subset=['Year/month','Entry Date','Document Type','Company Code','G/L Account',
                                     'Profit Center','Amount in local currency'],keep=False)

    df['Duplicate'] = dfDuplicates
    dfTempDups = df[df['Duplicate'] == True]
    dfTempDups.to_excel("outputDuplicates.xlsx", sheet_name='Sheet_name_1')
    print('process5 completed')

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
    
def processFile():
    #adding columns and setting defaults
    df['POtext'] = 'abc'
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

    totalRows1 = len(df.index)
    for row in df.index:
        poText = df['Text'].iloc[row]
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
        elif POhash in poText: #take left 13 or 14 for WBSText
            invText = poText[0:14]
            poText = 'PO' + poText[-8:]
        elif po in poText:
            invText = poText[0:14]
            poText = 'PO' + poText[-8:]
    
        df['POtext'].iloc[row] = poText
        df['WBSText'].iloc[row] = invText
        df['Abs'] = abs(df['Amount in local currency'])
        df['AddInverse'] = df['Amount in local currency'] * -1
  
    #overwrite WBSText with data from WBS column if not null
    dfWBS = df
    dfWBS['WBS element'].replace('', np.nan, inplace=True)
    #same as SQL Coalesce
    df['WBSText'] = np.where(df['WBS element'].isnull(),df['WBSText'],df['WBS element'])
    print('processing file completed, getting ready for comparing')

def processPamFile():
    dfPam = pd.read_excel(r'C:\Users\DawnBetzel\OneDrive - Warner Music Group\Projects\MES\Pam200920_091521.xlsx')
    print('getting ready to print dfPam')
    print(dfPam)

    
df = pd.read_excel(r'C:\Users\DawnBetzel\OneDrive - Warner Music Group\Projects\MES\Macro_ThreeYearDataSet_V3.xlsx')
#will have to do this for three files and append to the bottom of df
#MES can only export one year at a time??  Why is this need to know??

processFile()
process1()
process2()
process3()
process4('StringMatch1')
process4('StringMatch2')
process5()
process6()
process7()
processNum()
#processPamFile()
print('printing duplicate')
print(df['Duplicate'] == True)
print('printing matchid > 0')
print(df['MatchID'] > 0)
print('printing AlexCode')
print(df['AlexCode'] == True)



#fuzzy matching 25% inverse
dfSonopressTest = df['Vendor Description'].str.contains('SONOPRESS GMBH')
df['Sonopress'] = dfSonopressTest
dfDistinct = df[['Document Type','Company Code','G/L Account','Profit Center','Abs']]
dfDistinct['Abs'].drop_duplicates().sort_values()

df.to_excel("outputAll.xlsx", sheet_name='Sheet_name_1')
print('output Excel files completed')
