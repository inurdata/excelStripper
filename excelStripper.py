import csv
import xlrd
import sys

#global vars
inFile = ""
outFile = ""
keywordInput = ""
keys = [""]

#excelStripper.py by inurdata
#This program strips rows from csv files if they contain a keyword.
#If you import an excel file it converts it to csv.

while inFile = ""
	inFile = raw_input("Please enter csv file location ie C:\Path\input.csv: ")
	if inFile == "":
		print "please enter a file location"
	else:
		print "File location =", inFile
		
#Get keyword input from user
keywordInput = raw_input("Enter keywords separate by a comma ie key0, key1, keyETC...(Press Enter for none): ")
keys = keywordInput.split(",")
fileInput = raw_input("Please enter text file keyword list ie C:\Path\keys.txt (Press Enter for none): ")
#Check for blank input
if fileInput != "":
	fileInput = open(fileInput, 'r')
for line in fileInput.readlines():
	keys.append(line.rstrip('\n'))
#filter out empty strings or absent user input.
keys = filter(None, keys)
print "Keywords = ", keys

#ask for output location
outFile = raw_input("Please enter output file location ie C:\Path\output.csv: ")
print "Output file location = ", outFile

with open(inFile) as inp, open (outFile, 'w') as outp:
	print "working..."
	for line in inp:
		if not any(i in line for i in keys):
			outp.write(line)
inp.close()
outp.close()
print "done!"
exit()

