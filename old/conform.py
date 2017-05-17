# ******************************************************************************
# Author: Son D. Ngo
# Contributor: Steve Papaccio
# Updated: 11/9/2016

# Description: Script to unify institution's names using 100% matching
#	approach with a given dictionary. Data will be preprocessed by stripping
# 	leading and trailing white spaces, and words are case-insensitive when
# 	the matching happens.
# ******************************************************************************

# packages
import sys
import tkinter as tk
from tkinter import filedialog
from tkinter import simpledialog
from collections import defaultdict
import functions

def main():

	# STEP 1: Reading in command line and make necessary file conversion to csv 
	argc = len(sys.argv);

	# user specifies their own files through command line
	if (argc == 4):
		user_input = sys.argv[1]
		dictionary_input = sys.argv[2]
		column_name = sys.argv[3]

	# user uses the config (.ini) file
	elif (argc == 2):
		config_input = sys.argv[1]
		if (".ini" not in config_input):
			print("Invalid config file (.ini). Please try again!")
			quit()
		config = functions.configparser.ConfigParser()
		config._interpolation = functions.configparser.ExtendedInterpolation()
		config.read(config_input)
		user_input = config["DEFAULT"]["InputFile"]
		dictionary_input = config["DEFAULT"]["DictionaryFile"]
		column_name = config["DEFAULT"]["ColumnName"]

	# user chooses file with a dialog
	elif (argc == 1):
		root = tk.Tk()
		root.withdraw()
		user_input = filedialog.askopenfilename(title = "Choose file to conform")
		if (user_input == ""):
			print("Invalid input file. Please try again!")
			quit()
		dictionary_input = filedialog.askopenfilename()
		if (dictionary_input == ""):
			print("Invalid dictionary file. Please try again!")
			quit()
		# column_name = input("Please enter the column to conform: ")
		column_name = simpledialog.askstring(prompt = "Please enter the column name to conform: ", title="Message Box")

	# error message and exit with invalid command
	else:
		print("Invalid command. Please try again (refer to README.md)")
		quit()

	print("\n============= INPUTS INFORMATION =============")
	print("File to conform: " + user_input)
	print("Column to conform: " + column_name)
	print("Dictionary file: " + dictionary_input)
	print("==============================================\n")

	confirmation = input("Are the above information is correct? Would you like to proceed? (y or n): ")

	if (confirmation != "y"):
		print("Please modify your config (.ini) file or select the right file using the file dialog. Exiting...")
		quit()

	dictionary_input_csv = ""
	file_name = user_input[0:user_input.rfind(".")]
	file_format = ".xlsx"

	print(file_name)

	# converting steps
	if (user_input.endswith(".xlsx") or user_input.endswith(".xls")):
		user_input = functions.csv_from_excel(user_input)

	if (dictionary_input.endswith(".xlsx") or dictionary_input.endswith(".xls")):
		dictionary_input_csv = functions.csv_from_excel(dictionary_input)
		dictionary_input = dictionary_input_csv

	# variables for data structures
	columns = defaultdict(list)
	dictionary = defaultdict(list)
	mapper = {};
	no_match = [];
	match = [];
	header = [];

	# STEP 2: READ IN & SET UP DATA STRUCTURES

	# set up the user input map
	with open(user_input) as f:
		reader = functions.csv.DictReader(f)
		header = reader.fieldnames
		for row in reader:
			for (k,v) in row.items():
				columns[k].append(v)

	# set up the dictionary map
	with open(dictionary_input) as f:
		reader = functions.csv.DictReader(f)
		for row in reader:
			for (k,v) in row.items():
				dictionary[k].append(v)

	# create a mapper data structure using the dictionary
	for index in range(0,len(dictionary["AS ENTERED"])):
		mapper[(dictionary["AS ENTERED"])[index].lower()] = (dictionary["AS CLEANED"])[index]

	# STEP 3: MATCHING ALGORITHM & WRITE RESULTS TO FILE
	recipients = columns[column_name]
	index = 0;
	for item in recipients:

		# remove trailing, leading and multiple-in-text white spaces
		# then turn the name to lower case
		item = item.strip();
		item = " ".join(item.split())
		processedItem = item.lower();

		# start the matching
		if processedItem in mapper.keys():
			if not processedItem in match:
				match.append(processedItem)
			columns[column_name][index] = mapper[processedItem]
		else:
			if not item in no_match:
				no_match.append(item)
		index += 1

	# producing results
	functions.write_conform_result(columns, header, file_name + "_conformed" + file_format)
	functions.write_conform_result(no_match, [], file_name + "_non_matched" + file_format)
	# functions.write_conform_result(match, [], file_name + "_matched" + file_format)

	# prints
	print("\n=============SUMMARY RESULT=============")
	print("length of dictionary: " + str(len(mapper)))
	print("number of rows processed: " + str(len(recipients)))
	print("number of non_match (without duplicates): " + str(len(no_match)))
	print("number of matches (without duplicates): " + str(len(match)))
	print("==========================================\n")

	continueToUpdate = input("Conforming process has completed. Would you like to review and update non-matches now? (y or n): ");
	if (continueToUpdate == "y"):
		print("Running update.py")
		import update

	print("\nScript Finished.\n")

	# remove intermediary files
	if (dictionary_input_csv != ""):
		functions.os.remove(dictionary_input_csv)

if __name__ == "__main__":
	main()