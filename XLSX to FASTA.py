# -*- coding: utf-8 -*-
"""
Excel (.xlsx) to FASTA file converter.
Assumes the first row of the file is for titles
Exports data in the following format:
    >lcl|accession| organism name
    sequence
lcl is used in the FASTA format as an internal accession number, not linked to any database.
"""

import xlrd, re

#Import File
fileName = input("Enter the name of the file, including suffix:  ")
mainWorkbook = xlrd.open_workbook(fileName) #Opens the workbook
#Opens the first sheet of the file. This assumes that the data we care about is in the first sheet.
mainSheet = mainWorkbook.sheet_by_index(0)

def getColumnNum(questionText):
    userColumn = input(questionText)
    return int(ord(userColumn.lower())) - 97    #Converts the letter to the ASCII number for it, then substracts 97. A=97, and we want column A to be column 0.
   
#Imports first column, the accession numbers. 
accessionList = []
accessionColumn = getColumnNum("Enter the column letter for the Accession Numbers: ")
print(accessionColumn)
for i in range(1,mainSheet.nrows):
    accessionList.append(mainSheet.cell_value(i,accessionColumn))

#Imports second column, the amino acid sequence
sequenceList = []
sequenceColumn = getColumnNum("Enter the column letter for the sequence: ")
for i in range(1,mainSheet.nrows):
    sequenceList.append(mainSheet.cell_value(i,sequenceColumn))
    
protListRaw = []
protColumn = getColumnNum("Enter the column letter for the Protein ID: ")
for i in range(1,mainSheet.nrows):
    protListRaw.append(mainSheet.cell_value(i,protColumn))

#Sanitize protList
protList = protListRaw
'''for i in range(0, len(protListRaw)):
    tempProtList = []
    tempProtList.append(protListRaw[i].split(' '))
    print(tempProtList)
    if(len(tempProtList)==1):
        protList.append(tempProtList[0][0])
    else:
        protList.append(tempProtList[0][0]+tempProtList[0][1])
'''

#Imports eigth column, the organsim plus some extra info that we need to sanitize
organismListRaw = []
organismColumn = getColumnNum("Enter the column letter for the organism: ")
for i in range(1,mainSheet.nrows):
    organismListRaw.append(mainSheet.cell_value(i,organismColumn))

#Sanitizes organism list by only keeping the first two terms
organismList = []
for i in range(0,len(organismListRaw)):
    tempString = ''
    tempStringSplit = []
    tempString = organismListRaw[i]
    tempStringSplit = tempString.split('(')
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
outFile = open(newFileName+".FASTA","w+",encoding='utf-8')
#Writes file
for i in range(0,len(organismList)):
    outFile.write(">"+accessionList[i]+" | "+re.sub(r'\(.*\)','',protList[i])+" | "+organismList[i].replace(')','')+" | AMPs\n"+sequenceList[i]+"\n")
outFile.close() #Closes out the file
print("File written to "+newFileName+".FASTA successfully.")