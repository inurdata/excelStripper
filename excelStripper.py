import csv
import xlrd
import argparse

#usage info
def msg(name=None):
	return '''
	|  | |       _______|__________*_____________________________
	 \/  |      |_____  |   |/   \ | |/    \|/    \|    \ |/   
	 /\  |            | |   |      | |\____/|\____/|____/ |    
	|  | |____   ____ | |   |      | |      |      \_____ |     
	                                 |      |
	/  \      _______________________________________________________________________________________
    	|  |     |excelStripper deletes rows from a csv file based on keywords.                           \
    	@  @    / Usage: excelStripper.py -h -i INPUT -o OUTPUT -k KEYFILE -K KEYWORDS -s SHEETNAME -g -G  \   
    	|  |   /  >>>>if you don't specify an output file the output will be saved as output.csv/           |
    	|| |/ / >>>>if you input an .xlsx or .xls file and don't specify a sheet it defaults to its filename/
    	|| ||  *-------------------------------------------------------------------------------------------* 
    	|\_/|
    	\___/ 
	'''

#global vars
inFile = ""
outFile = ""
keywordInput = ""
fileInput = ""
sheet = "Sheet1"
workBook = ""
keys = []
parser = argparse.ArgumentParser(usage=msg())

#excelStripper.py by inurdata
#This program strips rows from csv files if they contain a keyword.
#If you import an excel file it converts it to csv.
#if you don't specify output it saves the csv as output.csv

#command line args
parser.add_argument('-i', '--input', type=str, help="input file location: C:\Path\input.csv")
parser.add_argument('-o', '--output', type=str, help="output file location: C:\Path\output.csv")
parser.add_argument('-k', '--keyfile', type=str, help="keyword text file: C:\Path\keys.txt")
parser.add_argument('-K', '--keywords', type=str, help="comma separated keywords: key0,key1,keyN")
parser.add_argument('-s', '--sheet', type=str, help="sheet to use if inputting an xlsx or xls file: SheetName")
parser.add_argument('-g', '--guided', help="use by itself for guided command line mode", action="store_true")
parser.add_argument('-G', '--gui', help="use by itself for GUI mode", action="store_true")
parser.parse_args()
args = parser.parse_args()

#FUNCTIONS
def getSheetName(inputFile):
	from os.path import basename
	s = basename(inputFile)
	s = s.rsplit(".",1)[0]
	return s;

def setOutFile(inputFile):
	from os.path import basename
	file = basename(inputFile)
	path = inputFile.replace(file,'')
	out = path+"output.csv"
	return out;

#MODES_______________________________________________________________________________________________________________________
#GUIDED MODE
if args.guided:
	while inFile == "":
		inFile = raw_input("Please enter csv file location ie C:\Path\input.csv: ")
		if inFile == "":
			print "Please enter a file location!"
		#check for xlsx or xls file and ask for sheet name
		if inFile.endswith(('.xlsx', '.xls')):
			sheet = raw_input("Enter your sheet name or press enter to skip (default is name of file): ")
			if sheet is "":
				sheet = getSheetName(inFile)
		#Get keyword input from user
		keywordInput = raw_input("Enter keywords separate by a comma ie key0, key1, keyETC...(Press Enter for none): ")
		fileInput = raw_input("Please enter text file keyword list ie C:\Path\keys.txt (Press Enter for none): ")

		#ask for output location
		outFile = raw_input("Please enter output file location ie C:\Path\output.csv: ")


#GUI MODE
elif args.gui:
	from Tkinter import *
	from Tkinter import messagebox
	import Tkinter, Tkconstants, tkFileDialog
	import os
	gui = Tk()
	cwd = os.getcwd()
	gui.resizable(width=True, height=True)
	gui.geometry('{}x{}'.format(400, 400))

	#Button Definitions
	def inputButton():
		gui.inFile = tkFileDialog.askopenfilename(initialdir = cwd, title = "Select INPUT file", filetypes = (("CSV Files","*.csv"),("Excel Files","*.xlsx"),("All Files", "*.*")) )
		global inFile
		global inLabel
		inFile = gui.inFile
		inLabel.config(text=gui.inFile)
		return;
	def outputButton():
		gui.outFile = tkFileDialog.asksaveasfile(initialdir = cwd, mode="w", title="Save OUTPUT file as", defaultextension=".csv", filetypes = (("CSV Files", "*.csv"),("All Files", "*.*")))
		global outFile
		global outLabel
		outFile = gui.outFile
		outLabel.config(text=gui.outFile)
		return;
	def keyButton():
		gui.fileInput = tkFileDialog.askopenfilename(initialdir = cwd, title = "Select KEY WORD file", filetypes = (("TXT Files","*.txt"),("All Files", "*.*")) )
		global fileInput
		global keyLabel
		fileInput = gui.fileInput
		keyLabel.config(text=gui.fileInput)
		return;
	def complete():
		global inFile
		global keywordInput
		global fileInput
		if inFile is "":
			messagebox.showerror("ERROR!", "You need to select an input file.")
			return;
		keywordInput = keyTxt.get()
		if keywordInput or fileInput is "":
			messagebox.showerror("ERROR!", "You need to enter a keyword or keyword file.")
			return;
		gui.destroy()
		return;
	def quit():
		gui.destroy()
		exit()
		return;

	#Buttons and things
	HEAD0=Label(gui, text="	|  | |       _______|______ *_____________________________")
	HEAD1=Label(gui, text="	\/  |      |_____   |   |/   \ | |/    \|/    \|    \ |/")
	HEAD2=Label(gui, text=" 	/\  |                | |   |       | |\__ /|\__ /|__ / |")
	HEAD3=Label(gui, text="	|  | |__________| |   |       | |       |         \__  |")
	HEAD4=Label(gui, text="	                                       |       |")
	HEAD0.pack(anchor=Tkinter.W)
	HEAD1.pack(anchor=Tkinter.W)
	HEAD2.pack(anchor=Tkinter.W)
	HEAD3.pack(anchor=Tkinter.W)
	HEAD4.pack(anchor=Tkinter.W)
	INFO = Label(gui, justify=LEFT, text="excelStripper deletes rows from a csv file based on keywords.")
	INFO.pack()
	inLabel = Label(gui, text=inFile)
	inLabel.pack()
	IN = Button(gui, text="Input File", command=inputButton, height=1, width=10)
	IN.pack()
	outLabel = Label(gui, text=outFile)
	outLabel.pack()
	OUT = Button(gui, text="Output File", command=outputButton, height=1, width=10)
	OUT.pack()
	keyLabel = Label(gui, text=fileInput)
	keyLabel.pack()
	KEY = Button(gui, text="Key File", command=keyButton, height=1, width=10)
	KEY.pack()
	txtLabel = Label(gui, text="Enter comma separated keywords here ie Key1, Key2, KeyETC: ")
	txtLabel.pack()
	keyTxt = Entry(gui)
	keyTxt.pack()
	keyTxt.focus_set()
	OK = Button(gui, text="GO!", command=complete, height=1, width=10)
	OK.pack()
	EXIT = Button(gui, text="Exit", command=quit, height=1, width=10)
	EXIT.pack()

	gui.mainloop()

#CMD LINE MODE
elif args.input:
	inFile = args.input
	if args.output:
		outFile = args.output
	if args.keywords:
		keywordInput = args.keywords
	if args.keyfile:
		fileInput = args.keyfile

#HELP
else:
	parser.print_help()
	exit()

#ERROR HANDLING & WORK PORTION__________________________________________________________________________________________

#infile error handling
if inFile is "":
	print "ERROR: you did not enter an input file."

# manual key input handler
if keywordInput != "":
	keys = keywordInput.split(",")

#file input handler
if fileInput.lower().endswith('.txt'):
	print "ERROR: your keyword file is not of the appropriate file type (.txt)"
if fileInput != "":
	fileInput = open(fileInput, 'r')
	for line in fileInput.readlines():
		keys.append(line.rstrip('\n'))

#filter out empty strings or absent user input.
keys = filter(None, keys)

#ERROR HANDLING
if not keys:
	print "ERROR: you didn't enter in any keys, key file, or your key file is empty"
	exit()
if not inFile.lower().endswith(('.xlsx', '.xls', '.csv')):
	print "ERROR: your input file is not of the appropriate file type (.xlsx, .xls, .csv)"
	exit()

#convert xlsx file or xls file to csv
if inFile.endswith('.xlsx') or inFile.endswith('.xls'):
	print "Converting xlsx/xls file to csv..."
	if args.sheet:
		sheet = args.sheet
	workBook = xlrd.open_workbook(inFile)
	workSheet = workBook.sheet_by_name(sheet)
	cFile = inFile+".csv"
	csvFile = open(cFile, 'wb')
	wr = csv.writer(csvFile, quoting=csv.QUOTE_ALL)
	for rowNum in xrange(workSheet.nrows):
		wr.writerow(list(x.encode('utf-8') if type(x) == type(u'') else x for x in workSheet.row_values(rowNum)))
	inFile = csvFile.name
	csvFile.close()
	print "Conversion Complete"

#set outfile if none set
if outFile is "":
	outFile = setOutFile(inFile)

#print out options
print "Input file = ", inFile
print "Output file = ", outFile
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

