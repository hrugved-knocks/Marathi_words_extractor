#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Jul 27 10:58:2 2019
Completed on Sat Jul 27 22:33:26 2019

@author: stableaf_
"""
#First things first, Importing required Libraries

import pdftotext
import re
from xlwt import Workbook
wb = Workbook()
sheet1 = wb.add_sheet('Sheet 1') 

#Declaring the Variables

text = ""
wordlist = []
global_wordlist = []
i=0
k=1
j=0
count=0
punctuations = '''‘’!()-[]{};:'"\,<>./?@#$%^&*_~\n'''

#Accessing the File to process

with open("/home/stableaf_/Desktop/Marathi.pdf", "rb") as f:
    pdf = pdftotext.PDF(f)
    
#Processing the file
    
for page in pdf:

#Declaring the Temp Variable used for each page
    
    text = page
    text1= ""
    wordfreq = []
    p=0
    
#Processing the Text
    
    for char in text:
        if char not in punctuations:
            text1 = text1 + char
            
    sentences = re.split(' ',text1)
    wordlist = list(sentences)
    #print(wordlist)
    global_wordlist.extend(wordlist)
    #print("\n",global_wordlist)
    
#counting the Frequency of each word
        
for w in global_wordlist:
    wordfreq.append(global_wordlist.count(w))
temp = int(len(wordfreq)) 

#Writing information to workbook
 
for _ in range(temp):
    if p <= temp:
        sheet1.write(i, j, global_wordlist[p])
        sheet1.write(i, k, wordfreq[p])
        i+=1
        p+=1

    else:
        break
    
#Closing workbook  
        
wb.save('xlwt example.xls') 
print("\n File saved Successfully!!")
