# Code by Josh Kuan
# Sample cmd: python xlsExtractor.py filename.xls

import re
import sys
import xlrd

def cellTranslate(topLeft):
	# Quality of Life code: Translates alphanumeric Excel cell coordinate to index tuple
	# Translates column from alphabetical to index
	regexObj = re.search("[A-Z]*", topLeft)
	if not regexObj.group(0):
		print("No column specified")
		return
	idxTopLeft = [0,-1]
	colTopleft = regexObj.group(0)
	for i in range(len(colTopleft)):
		idxTopLeft[1] += (26**i) * (ord(colTopleft[len(colTopleft)-1-i]) % 64)

	# Translates row number
	idxTopLeft[0] = int(re.sub(colTopleft, "", topLeft))-1

	return idxTopLeft

def tableExtract(sheet, outFile, topLeft, rowCount, colCount):
	print(outFile+" START...", end="")
	with open(outFile, 'w') as w:
		for i in range(topLeft[0], topLeft[0]+rowCount):
			currentRowStr = ""
			for j in range(topLeft[1], topLeft[1]+colCount):
				currentRowStr += str(sheet.cell_value(i, j))
				if j < topLeft[1]+colCount-1:
					currentRowStr += ","
			w.write(currentRowStr+"\n")
	w.closed
	print(" COMPLETE")
	return

if __name__ == "__main__":
	if len(sys.argv) < 2:
		sys.exit("Error: Missing input filename.")

	inFile = sys.argv[1]
	print("Input File:", inFile, "\n")

	# loc = ".\\filename.xls"
	loc = sys.argv[1]

	wb = xlrd.open_workbook(loc)
	sheet = wb.sheet_by_index(0)

	dictionary = {
		# "outputFilename.csv": ["topLeft", rowCount, coLCount]
		"table1.csv": ["C7", 17, 5],
		"table2.csv": ["B28", 14, 7],
		"table3.csv": ["J28", 9, 8],
		"table4.csv": ["B43", 14, 3],
		"table5.csv": ["J42", 9, 8]
	}

	for key in dictionary:
		value = dictionary[key]
		topLeft = cellTranslate(value[0])
		tableExtract(sheet, key, topLeft, value[1], value[2])

'''
Install xlrd package:
	-> Download tar.gz from https://pypi.org/project/xlrd/#files
	-> Extract and run `python setup.py install` from extracted setup.py directory

Notes:
	ord("A") = 65
	ord("Z") = 90

	WQ	= (26 * 23) + (17)
		= 615
	XFD	= (26^2 * 24) + (26 * 6) + (4)
		= 16384

	Time:
	((((0.1624537037037037 * 24) % 3) * 60) % 53) * 60
	0.1624537037037037 -> 03:53:56
'''