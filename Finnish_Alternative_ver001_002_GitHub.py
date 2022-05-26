# -*- coding: utf-8 -*-
"""
Created on Sun May 15 11:38:05 2022

@author: Nasim
"""

#Imports
print('Importing libraries...')
from datetime import datetime
import pandas as pd
import numpy as np
import random
import xlsxwriter

import matplotlib.pyplot as plt
#%matplotlib inline

print('Libraries imported!')
start_time = datetime.now()  #Timer start
print('Start time: ', start_time)

print('Checking parameters...')
#Ons and Offs

Windows_version = 1 #Run excel datafile on Windows  =1, or MacOS =0 (edited excel version for better usability)

#Load database file(s)
print('Loading database...')
#Sources:   #https://kaino.kotus.fi/sanat/nykysuomi/
            #https://metashare.csc.fi/repository/browse/finnish-verbal-colorative-constructions/a5d4be8c914c11e7ac29005056be118e7a5d0f46297b4d0e851bf10067440eb6/
        
# Read Excel file: Finnish_wordlist_edited_v003.xlsx
if Windows_version ==1: #for Windows 
    xls = pd.ExcelFile(r'D:\ ... \Finnish_wordlist_edited_v003.xlsx') # Insert your file directory
else: #for Mac
    xls = pd.ExcelFile(r'/Users/ ... /Finnish_wordlist_edited_v003.xlsx') # Insert your file directory
    
#timer
mean_time = datetime.now() #Timer end
print('Duration: {}'.format(mean_time - start_time))
print('database loaded!')

print('Creating dataframes...')
#Load sheets as dataframes
df1 = pd.read_excel(xls, 'WordsList') #Main Database

print('Dataframes created')
                   
#Progress timers
timer_percentages_interval_5s = np.arange(0.05,1.05,0.05)
timer_percentages_5s = [int(round(i * len(df1), 0)) for i in timer_percentages_interval_5s]
timer_percentages_interval_10s = np.arange(0.1,1.1,0.1)
timer_percentages_10s = [int(round(i * len(df1), 0)) for i in timer_percentages_interval_10s]
                   
print('Looping words and counting letters...')

#Create Excel File and sheet
workbook_FA  = xlsxwriter.Workbook('Finnish_Alphabet.xlsx')
worksheet_FALetters = workbook_FA.add_worksheet()

#Write headers in first row
worksheet_FALetters.write(0, 0, 'Letter')
worksheet_FALetters.write(0, 1, 'Letter Hits')
worksheet_FALetters.write(0, 2, 'Double Letter')
worksheet_FALetters.write(0, 3, 'Double Letter Hits')

List_of_letters = []
List_of_letters_hits = []
List_of_double_letters = []
List_of_double_letters_hits = []

count1 = 0
count2 = 0
count_Letters = 0
count_timer01 = 0
for i in range(1, len(df1) +1): #Loop over all words
#for i in range(1, 8): #Loop over the first few words
    #Progress timer for every 10%
    count_timer01 += 1
    if count_timer01 in timer_percentages_10s:
        if count_timer01 == timer_percentages_10s[0]:
            print('10%...')
            mean_time = datetime.now() #Timer end
            print('Duration: {}'.format(mean_time - start_time))
        if count_timer01 == timer_percentages_10s[1]:
            print('20%...')
            mean_time = datetime.now() #Timer end
            print('Duration: {}'.format(mean_time - start_time))
        if count_timer01 == timer_percentages_10s[2]:
            print('30%...')
            mean_time = datetime.now() #Timer end
            print('Duration: {}'.format(mean_time - start_time))
        if count_timer01 == timer_percentages_10s[3]:
            print('40%...')
            mean_time = datetime.now() #Timer end
            print('Duration: {}'.format(mean_time - start_time))
        if count_timer01 == timer_percentages_10s[4]:
            print('50%...')
            mean_time = datetime.now() #Timer end
            print('Duration: {}'.format(mean_time - start_time))
        if count_timer01 == timer_percentages_10s[5]:
            print('60%...')
            mean_time = datetime.now() #Timer end
            print('Duration: {}'.format(mean_time - start_time))
        if count_timer01 == timer_percentages_10s[6]:
            print('70%...')
            mean_time = datetime.now() #Timer end
            print('Duration: {}'.format(mean_time - start_time))
        if count_timer01 == timer_percentages_10s[7]:
            print('80%...')
            mean_time = datetime.now() #Timer end
            print('Duration: {}'.format(mean_time - start_time))
        if count_timer01 == timer_percentages_10s[8]:
            print('90%...')
            mean_time = datetime.now() #Timer end
            print('Duration: {}'.format(mean_time - start_time))
        if count_timer01 == timer_percentages_10s[9]:
            print('100%...')
            mean_time = datetime.now() #Timer end
            print('Duration: {}'.format(mean_time - start_time))
    
    count1 += 1
    #print(df1.iloc[count1-1,:][1])
    for ii in range(1, df1.iloc[i-1,:][4] + 1): # Loop over number of letters in word
        count2 += 1
        #print(df1.iloc[count1-1,:][1][ii-1])
        if df1.iloc[count1-1,:][1][ii-1] not in List_of_letters: #Unique Letters
            List_of_letters.append(df1.iloc[count1-1,:][1][ii-1])
            List_of_letters.index(df1.iloc[count1-1,:][1][ii-1])
            # count letter occurence
            List_of_letters_hits.append(1)
            #Fill 1st column with unique Letters and/or characters
            worksheet_FALetters.write(List_of_letters.index(df1.iloc[count1-1,:][1][ii-1]) +1, 0, df1.iloc[count1-1,:][1][ii-1])
            #Fill 2nd column with 1 if first occurence
            worksheet_FALetters.write(List_of_letters.index(df1.iloc[count1-1,:][1][ii-1]) +1, 1, 1)
        else:
            List_of_letters_hits[List_of_letters.index(df1.iloc[count1-1,:][1][ii-1])] += 1
            #Fill 2nd column with counted hit number for the specific letter and/or character
            worksheet_FALetters.write(List_of_letters.index(df1.iloc[count1-1,:][1][ii-1]) +1, 1, List_of_letters_hits[List_of_letters.index(df1.iloc[count1-1,:][1][ii-1])])
            
        if ii > 0: # analyze double letters occurences
            if df1.iloc[count1-1,:][1][ii-1] == df1.iloc[count1-1,:][1][ii-2]:
                if df1.iloc[count1-1,:][1][ii-1] not in List_of_double_letters:
                    List_of_double_letters.append(df1.iloc[count1-1,:][1][ii-1])
                    #Add 1 to the hit list
                    List_of_double_letters_hits.append(1)
                    #Fill 3rd column with unique double letters and/or characters
                    worksheet_FALetters.write(List_of_double_letters.index(df1.iloc[count1-1,:][1][ii-1]) +1, 2, df1.iloc[count1-1,:][1][ii-1])
                    #Fill 4th column with 1 if first occurence
                    worksheet_FALetters.write(List_of_double_letters.index(df1.iloc[count1-1,:][1][ii-1]) +1, 3, 1)
                else:
                    List_of_double_letters_hits[List_of_double_letters.index(df1.iloc[count1-1,:][1][ii-1])] += 1
                    #Fill 4th column with counted hit number for the specific double letter and/or character
                    worksheet_FALetters.write(List_of_double_letters.index(df1.iloc[count1-1,:][1][ii-1]) +1, 3, List_of_double_letters_hits[List_of_double_letters.index(df1.iloc[count1-1,:][1][ii-1])])

workbook_FA.close()            

#timer
end_time = datetime.now() #Timer end
print('Duration: {}'.format(end_time - start_time))

print('End script!')