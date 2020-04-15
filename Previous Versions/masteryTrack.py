## Importing Packages to run Pandas

import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import os
import math
import re

from sklearn.preprocessing import OneHotEncoder
from sklearn import metrics
from sklearn.datasets import make_classification
from sklearn.linear_model import LogisticRegression
from sklearn.linear_model import LinearRegression
from sklearn.ensemble import (RandomTreesEmbedding, RandomForestClassifier,
                              GradientBoostingClassifier)
from sklearn.ensemble import RandomForestRegressor
from sklearn.preprocessing import OneHotEncoder
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import Imputer
from sklearn.model_selection import cross_val_score
from sklearn.svm import SVR
import glob

## Import Standards for Course

path = r'.\Standards'
filename = glob.glob(path + "/*.xlsx")

StandardsList = pd.read_excel(filename[0])

SLColumns = StandardsList.columns.values
Codes = StandardsList.loc[:,'Standard Code']
Description = StandardsList.loc[:,'Description']
Unit = StandardsList.loc[:,'Unit']
Priority = StandardsList.loc[:,'Priority']
DF_temp = pd.DataFrame(index = StandardsList.index.values, columns = ['Last Date Assessed'])

StandardsList = pd.concat([Codes, Description, Unit, Priority, DF_temp], axis=1)

StandardsList.head()

## Determine date of each assessment

path = r'.\TestInfo'
filename = glob.glob(path + "/*.xlsx")

DatesAssessed = []
for test in range(len(filename)):
    DatesAssessed.append(filename[test][len(path)+1:len(path)+9])
    
#Processing the dates into better format
DatesAssessed = [e[0:4] + "-" + e[4:6] + "-" + e[6:8] for e in DatesAssessed]

## Import Student Roster

path = r'.\Roster'
filename = glob.glob(path + "/*.xlsx")

Roster = pd.read_excel(filename[0])

Roster.index = Roster.loc[:,'Student ID']
Roster = Roster.sort_index()
Roster = Roster.drop(['Student ID'],axis = 1)
Roster.head()

## Define function to clean test info page

def cleanTestInfo(TestInfoFileName):
    TestInfo = pd.read_excel(TestInfoFileName, skiprows=9)
    TestInfo.index = range(1, len(TestInfo.index)+1)
    Standards = TestInfo.loc[:,'(Primary) Standard']
    Type = TestInfo.loc[:,'MC, OER (Question Group)']
    Points = TestInfo.loc[:,'Possible Points']
    Correct = TestInfo.loc[:,'Correct Answer']

    TestInfo = pd.concat([Standards, Type, Points, Correct], axis = 1)
    Qseries = TestInfo.loc[:,'(Primary) Standard']
    LastQuestion = Qseries.last_valid_index()
    TestInfo = TestInfo.loc[1.0:LastQuestion,:]
    return TestInfo
	
## Define function to clean response matrices

def cleanResponses(ResponsesFileName,TestInfo):
    Responses = pd.read_excel(ResponsesFileName)
    Questions = Responses.columns.values[9:]
    StudentID = Responses.loc[:,['Local Student Id']]
    StudentResponses = Responses.loc[:,Questions]
    Responses = pd.concat([StudentID, StudentResponses], axis = 1)
    Responses.index = Responses.loc[:,'Local Student Id']
    Responses = Responses.drop(['Local Student Id'], axis = 1)
    Responses.columns = TestInfo.index
    Responses = Responses.sort_index()
    return Responses
	
def createBinary(Responses, TestInfo):
    BinaryMatrix = Responses.copy()
    for question in TestInfo.index.values:
        if TestInfo.loc[question,'MC, OER (Question Group)'] == 'MC':
            for student in Responses.index.values:
                if TestInfo.loc[question,'Correct Answer'] == Responses.loc[student,question]:
                    BinaryMatrix.loc[student,question] = TestInfo.loc[question,'Possible Points']
                else:
                    BinaryMatrix.loc[student,question] = 0
        else:
            continue
    return BinaryMatrix
	
## Define function to calculate points per standard for a given test

def calcPPS(StandardsList,TestInfo):

    StandardIDs = StandardsList.loc[:,'Standard Code']
    PPS = pd.DataFrame(index = StandardIDs, columns = ['Points'])
    for standard in StandardIDs:
        PPS.loc[standard,'Points'] = 0
        for question in TestInfo.index.values:
            if TestInfo.loc[question,'(Primary) Standard'] == standard:
                PPS.loc[standard,'Points'] = PPS.loc[standard,'Points'] + TestInfo.loc[question,'Possible Points']
            else:
                continue
    return PPS
	
## Define function to calculate standards matrix per student

def createStandardsMatrix(BinaryMatrix, TestInfo, PPS):

    StandardIDs = PPS.index.values
    StandardsbyStudent = pd.DataFrame(index = BinaryMatrix.index.values, columns = StandardIDs, data = None)
    
    for standard in StandardsbyStudent.columns:
        AlignedQuestions = list()
        for question in TestInfo.index.values:
            if TestInfo.loc[question,'(Primary) Standard'] == standard:
                AlignedQuestions.append(question)
            else:
                continue
        for student in StandardsbyStudent.index:
            if PPS.loc[standard,'Points'] == 0:
                StandardsbyStudent.loc[student,standard] = 0
            else:
                PointsEarned = BinaryMatrix.loc[student,AlignedQuestions].sum()
                StandardsbyStudent.loc[student,standard] = (PointsEarned)
                
    return StandardsbyStudent
	
## Read Filenames for Test Info Pages and Response Matrices
# Use this later when working with multiple tests at once

path1 = r'.\TestInfo'
TestInfoNames = glob.glob(path1 + "/*.xlsx")

path2 = r'.\Responses'
ResponsesNames = glob.glob(path2 + "/*.xls")

## Import/Clean Info Pages, Import/Clean Responses, Create Binary Matrices

TestInfos = []
Responses = []
Binaries = []
PPSs = []
StandardsMatrices = []

for testName in TestInfoNames:
    DF_temp = cleanTestInfo(testName)
    TestInfos.append(DF_temp)
    
count = 0
for responsesName in ResponsesNames:
    DF_temp = cleanResponses(responsesName,TestInfos[count])
    Responses.append(DF_temp)
    count = count + 1
    
for testNum in range(len(Responses)):
    DF_temp = createBinary(Responses[testNum],TestInfos[testNum])
    Binaries.append(DF_temp)
    DF_temp = calcPPS(StandardsList, TestInfos[testNum])
    PPSs.append(DF_temp)
    DF_temp = createStandardsMatrix(Binaries[testNum],TestInfos[testNum],PPSs[testNum])
    StandardsMatrices.append(DF_temp)
    
PPSs[0].head()

## Add the last date of assessment to Standards list

for test in range(len(PPSs)):
    print('test', test)
    count = 0
    for standard in PPSs[test].index.values:
        print(PPSs[test].loc[standard,'Points'], count)
        if PPSs[test].loc[standard,'Points'] == 0:
            StandardsList.loc[count, 'Last Date Assessed'] = StandardsList.loc[count, 'Last Date Assessed']
        else:
            StandardsList.loc[count, 'Last Date Assessed'] = DatesAssessed[test]
        count = count + 1
        
StandardsList

## Sum Student Mastery Matrices and PPSs

StandardsMatricesSUM = sum(StandardsMatrices)
PPSsSUM = sum(PPSs)

OverallMastery = StandardsMatricesSUM.copy()

for student in StandardsMatricesSUM.index:
    for standard in PPSsSUM.index:
        if PPSsSUM.loc[standard,'Points'] == 0:
            OverallMastery.loc[student,standard] = 'NaN'
        else:
            OverallMastery.loc[student,standard] = StandardsMatricesSUM.loc[student,standard]/PPSsSUM.loc[standard,'Points']
            #Depreciated version with *100 for percent
            #OverallMastery.loc[student,standard] = StandardsMatricesSUM.loc[student,standard]/PPSsSUM.loc[standard,'Points']*100

OverallMastery

#Get Roster df in order
Roster_processed = Roster.copy()
#cols = Roster2.columns.tolist() 
#cols = [cols[:0]]+[cols[2]]+[cols[1]]
#Roster2 = Roster2[cols]
#Roster2 = Roster2.ix[:, cols]
Roster_processed = Roster_processed[['Last, First', 'Section', 'Teacher']]
# 	Last, First 	Teacher 	Section

Roster_processed.head()

StandardsList_processed = StandardsList.copy()
StandardsList_processed = StandardsList_processed [['Unit', 'Last Date Assessed', 'Priority', 'Standard Code']]
StandardsList_processed = StandardsList_processed.fillna(value='')
StandardsList_processed = StandardsList_processed.T
StandardsList_processed.head()

#Time to Upload!
from df2gspread import df2gspread as d2g
import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

credentials2 = ServiceAccountCredentials.from_json_keyfile_name('creds.json', scope)

gc = gspread.authorize(credentials2)
# Create empty dataframe
df = pd.DataFrame()

# Create a column
#df['name'] = ['John2', 'Steve', 'Sarah']
#df.append
# use full path to spreadsheet file
#spreadsheet = '/some/folder/New Spreadsheet'
# or spreadsheet file id
spreadsheet = '1AvN6e8bUKdXq43lYTU0fnK4dmUjC5iRdYaOpaoMFLHc'
wks = 'Course Template'

#Upload Roster Data
d2g.upload(Roster_processed, gfile=spreadsheet, wks_name=wks, start_cell='A40', credentials=credentials2, clean=False, df_size=False, col_names=False, row_names=False)
#d2g.upload(df, gfile=spreadsheet, wks_name=wks, start_cell='A40', credentials=credentials2, clean=False, df_size=False, col_names=False, row_names=False

#Upload Standards Mastery Numbers
d2g.upload(OverallMastery, gfile=spreadsheet, wks_name=wks, start_cell='E40', credentials=credentials2, clean=False, df_size=False, col_names=False, row_names=False)

#Upload the Standards themselves to TWO locations.
d2g.upload(StandardsList_processed, gfile=spreadsheet, wks_name=wks, start_cell='E19', credentials=credentials2, clean=False, df_size=False, col_names=False, row_names=False)
d2g.upload(StandardsList_processed, gfile=spreadsheet, wks_name=wks, start_cell='E36', credentials=credentials2, clean=False, df_size=False, col_names=False, row_names=False)
