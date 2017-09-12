import csv
import xlrd
import argparse
import sys

#global vars
inFile = ""
outFile = ""
keywordInput = ""
fileInput = ""
keys = []
parser = argparse.ArgumentParser()

#excelStripper.py by inurdata
#This program strips rows from csv files if they contain a keyword.
#If you import an excel file it converts it to csv.
#if you don't specify output it saves the csv as .csv.new

#command line args
parser.add_argument('-i', '--input', type=str, help="input file location: C:\Path\input.csv")
parser.add_argument('-o', '--output', type=str, help="output file location: C:\Path\output.csv")
parser.add_argument('-k', '--keyfile', type=str, help="keyword text file: C:\Path\keys.txt")
parser.add_argument('-K', '--keywords', type=str, help="comma separated keywords: key0,key1,keyN")
parser.add_argument('-g', '--guided', help="use by itself for guided command line mode", action="store_true")
parser.add_argument('-G', '--gui', help="use by itself for GUI mode", action="store_true")
parser.parse_args()
args = parser.parse_args()

#_______________________________________________________________________________________________________________________
#GUIDED MODE
if args.guided:
	while inFile == "":
		inFile = raw_input("Please enter csv file location ie C:\Path\input.csv: ")
		if inFile == "":
			print "please enter a file location"
		else:
			print "File location =", inFile
		#Get keyword input from user
		keywordInput = raw_input("Enter keywords separate by a comma ie key0, key1, keyETC...(Press Enter for none): ")
		fileInput = raw_input("Please enter text file keyword list ie C:\Path\keys.txt (Press Enter for none): ")

		#ask for output location
		outFile = raw_input("Please enter output file location ie C:\Path\output.csv: ")
		if outFile == "":
			outFile = inFile+".new"
			print "Output file location = ", outFile

#GUI MODE
elif args.gui:
	print "in progress"
	exit()

#CMD LINE MODE
elif args.input:
	inFile = args.input
	if args.output:
		outFile = args.output
	else:
		outFile = inFile+".new"
	if args.keywords:
		keywordInput = args.keywords
	if args.keyfile:
		fileInput = args.keyfile

#HELP
else:
	parser.print_help()
	exit()
#______________________________________________________________________________________________________________________

# manual key input handler
if keywordInput != "":
	keys = keywordInput.split(",")

#file input handler
if fileInput != "":
	fileInput = open(fileInput, 'r')
	for line in fileInput.readlines():
		keys.append(line.rstrip('\n'))

#filter out empty strings or absent user input.
keys = filter(None, keys)
if not keys:
	print "ERROR: you didn't enter in any keys, key file, or your key file is empty"
	exit()

#print out options
print "input file = ", inFile
print "output file = ", outFile
print "Keywords = ", keys

#delete rows
with open(inFile) as inp, open (outFile, 'w') as outp:
	print "working..."
	for line in inp:
		if not any(i in line for i in keys):
			outp.write(line)
inp.close()
outp.close()
print "done!"
exit()

