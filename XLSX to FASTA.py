# -*- coding: utf-8 -*-
"""
Created on Thu Jan 30 15:37:18 2020

@author: tsilv
"""

import xlrd, time

#Import File
fileName = input("Enter the name of the file, including suffix:  ")
mainWorkbook = xlrd.open_workbook(fileName)
#Opens the first sheet of the file. This assumes that the data we care about is in the first sheet.
mainSheet = mainWorkbook.sheet_by_index(0)
#Sets cursor to 0,0
mainSheet.cell_value(0,0)


#Imports first column, the accession numbers. 
accessionList = []
for i in range(1,mainSheet.nrows):
    accessionList.append(mainSheet.cell_value(i,0))

#Imports second column, the amino acid sequence
sequenceList = []
for i in range(1,mainSheet.nrows):
    sequenceList.append(mainSheet.cell_value(i,1))

#Imports eigth column, the organsim plus some extra info that we need to sanitize
organismListRaw = []
for i in range(1,mainSheet.nrows):
    organismListRaw.append(mainSheet.cell_value(i,7))

#Sanitizes organism list by only keeping the first two terms
organismList = []
for i in range(0,len(organismListRaw)):
    tempString = ''
    tempStringSplit = []
    tempString = organismListRaw[i]
    tempStringSplit = tempString.split()
    #Some organism names aren't at least 2 words long
    if(len(tempStringSplit)==0):
        organismList.append("")
    elif(len(tempStringSplit)==1):
        organismList.append(tempStringSplit[0])
    else:
        organismList.append(tempStringSplit[0] + " " + tempStringSplit[1])

#Prepares new file
newFileName = input("Enter new, unique file name: ")
#Initializes the new file with read/write permission
outFile = open(newFileName+".FASTA","w+")
#Writes file
for i in range(0,len(organismList)):
    outFile.write(">lcl|"+accessionList[i]+"| "+organismList[i]+"\n"+sequenceList[i]+"\n")
outFile.close()

time.sleep(10)