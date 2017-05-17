# ******************************************************************************
# Author: Son D. Ngo
# Contributor: Steve Papaccio
# Updated: 11/9/2016

# Description: Script to allow user to update the dictonary by manually typing
#	in the conformed name for unmatched institutions. The updated dictionary
#	will be written to the same directory with the _udpated suffix to its name.
#	The file for unmatched names will also be updated. 

# The script currently only supports dictionary of type xlsx and xls
# ******************************************************************************

# packages
import sys
import tkinter as tk
from tkinter import filedialog
from collections import defaultdict
import functions 

def main():

	# STEP 1: Reading in command line and make necessary file conversion to csv

	argc = len(sys.argv)

	# user specifies their own files through command line
	if (argc == 3):
		user_input = sys.argv[1]
		dictionary_input = sys.argv[2]

	# user uses the config (.ini) file
	elif (argc == 2):
		config_input = sys.argv[1]
		if (".ini" not in config_input):
			print("Invalid config file (.ini). Please try again!")
			quit()
		config = functions.configparser.ConfigParser()
		config._interpolation = functions.configparser.ExtendedInterpolation()
		config.read(config_input)
		user_input = config["DEFAULT"]["Nonmatched"]
		dictionary_input = config["DEFAULT"]["DictionaryFile"]

	# user chooses file with a dialog
	elif (argc == 1):
		root = tk.Tk()
		root.withdraw()
		user_input = filedialog.askopenfilename()
		if (user_input == ""):
			print("Invalid input file. Please try again!")
			quit()
		dictionary_input = filedialog.askopenfilename()
		if (dictionary_input == ""):
			print("Invalid dictionary file. Please try again!")
			quit()

	# error message and exit with invalid command
	else:
		print("Invalid command. Please try again (refer to README.md)")
		quit()

	user_input_csv = ""

	# converting steps
	if user_input.endswith(".xlsx") or user_input.endswith(".xls"):
		user_input_csv = functions.csv_from_excel(user_input)

	if not dictionary_input.endswith(".xlsx"):
		dictionary_input = functions.xlsx_from_xls(dictionary_input)

	dictionary_input_csv = functions.csv_from_excel(dictionary_input)

	columns = defaultdict(list)
	dictionary = defaultdict(list)
	mapper = {};

	# STEP 2: READ IN & SET UP DATA STRUCTURES

	# reading and setting up data structures step
	with open(user_input_csv) as f:
		reader = functions.csv.DictReader(f)
		header = reader.fieldnames
		for row in reader:
			for (k,v) in row.items():
				columns[k].append(v)		

	with open(dictionary_input_csv) as f:
		reader = functions.csv.DictReader(f)
		for row in reader:
			for (k,v) in row.items():
				dictionary[k].append(v)

	# create a mapper data structure using the dictionary
	for index in range(0,len(dictionary["AS ENTERED"])):
		mapper[(dictionary["AS ENTERED"])[index].lower()] = (dictionary["AS CLEANED"])[index]

	book = functions.load_workbook(dictionary_input)
	sheet = book.get_sheet_by_name(book.get_sheet_names()[0])
	sheet.cell(row = 1, column = 3).value = "DATE UPDATED"

	# STEP 3: PROMPT USER INPUT AND SAVE TO UPDATED DICTIONARY
	index = 0;
	removedNonMatch = [];

	for item in columns["AS ENTERED"]:
		conformedName = columns["AS CLEANED"][index]
		if conformedName:
			print(conformedName)

		# remove trailing, leading and multiple-in-text white spaces
		# then turn the name to lower case
		item = item.strip();
		item = " ".join(item.split())
		processedItem = item.lower();

		# finish updating
		if not conformedName:
			# if removedNonMatch:
			# 	savedPath = dictionary_input[0:dictionary_input.find(".xlsx")] + "_updated.xlsx"
			# 	print("Save updated dictionary to: " + savedPath)
			# 	book.save(savedPath)
			# break
			index += 1
			continue

		# option to skip updating a particular entry
		elif processedItem in mapper.keys():
			if conformedName != mapper[processedItem]:
				print("Found a conflict between user and dictionary")
				print("AS ENTERED: " + item)
				print("AS CLEANED (user): " + conformedName)
				print("AS CLEANED (dictionary): " + mapper[processedItem])
				print("Which one would you like to keep? (0 for user, 1 for dictionary, or type a new name")
				message = input()
				if message == "1":
					conformedName = mapper[processedItem]
				elif message != "0":
					conformedName = message
		
		#  add to temporary worksheet
		removedNonMatch.append(item)
		row_to_write = sheet.max_row + 1
		sheet.cell(row = row_to_write, column = 1).value = item
		sheet.cell(row = row_to_write, column = 2).value = conformedName
		sheet.cell(row = row_to_write, column = 3).value = functions.datetime.date.today()
		print("Add " + "\"" + item + "\"" + " to dictionary");
		index += 1

	print("number of non-match removed: " + str(len(removedNonMatch)))

	# write merged dictionary to a new file if there is an update
	if removedNonMatch:
		savedPath = dictionary_input[0:dictionary_input.find(".xlsx")] + "_update_" + functions.datetime.date.today().strftime('%m-%d-%Y') + ".xlsx"
		print("Save updated dictionary to: " + savedPath)
		book.save(savedPath)

	# update non-match file (remove the ones that have been updated)
	columns["AS ENTERED"] = set(columns["AS ENTERED"]) - set(removedNonMatch)
	functions.os.remove(user_input)
	functions.write_update_result(columns["AS ENTERED"], user_input)

	# remove intermediary file
	if (user_input_csv != ""):
		functions.os.remove(user_input_csv)

	if (dictionary_input_csv != ""):
		functions.os.remove(dictionary_input_csv)

if __name__ == "__main__":
	main()