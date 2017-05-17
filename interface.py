# ******************************************************************************
# Author: Son D. Ngo
# Filename: interface.py
# Contributor: Steve Papaccio
# Updated: 5/17/2017

# Description: Script that implements the interface for the name conforming
# 				program. All the buttons, labels, input entries are created
# 				in this script, and hooked up to the right functions
# 				implemented in the functions.py script. 
# ******************************************************************************

# Modules & Packages
import functions
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

# the interface class
class App(tk.Frame):

	columnsHeader = ['(empty)']

	def __init__(self, master=None):
		super().__init__(master)
		self.initialize()

	# initialize the interface by constructing a grid
	# each item (label, input entry, button) belongs to a cell in this grid
	# items will be added by finishing one row first, then moving on to the next row
	# connect the button to the right callback function
	def initialize(self):
		self.grid()

		entryWidth = 50

		# string var to keep track of each input entry
		self.inputVarConform = tk.StringVar()	# input file used in conform
		self.inputVarMerge = tk.StringVar()		# input file used in merge
		self.dictVarConform = tk.StringVar()	# dictionary used in conform
		self.dictVarMerge = tk.StringVar()		# dictionary used in merge
		self.columnVar = tk.StringVar()			# which column to conform
		self.radioVar = tk.IntVar()				# whether to use full or abbreviated name

		currentRow = 0;

		## ROW 0 ##

		# input file input entry label
		inputLabel = tk.Label(self, anchor="w", fg="black", text="File to conform: ")
		inputLabel.grid(column=0, row=currentRow, sticky="EW")

		# input file input entry
		self.fileEntry = tk.Entry(self, textvariable = self.inputVarConform, width = entryWidth)
		self.fileEntry.grid(column=1, row=currentRow, sticky="EW")

		# input file browse button
		inputFileBrowseButton = tk.Button(self, text="Browse...", command=self.OnBrowseInputButtonClick)
		inputFileBrowseButton.grid(column=2, row=currentRow)

		## ROW 1 ##

		currentRow += 1

		# column drop-down menu label
		columnName = tk.Label(self, anchor="w", fg="black", text="Column to conform: ")
		columnName.grid(column=0, row=currentRow, sticky="EW")

		# column drop-down menu
		self.columnMenu = tk.OptionMenu(self, self.columnVar, *(self.columnsHeader))
		self.columnMenu.grid(column=1, row=currentRow, sticky="EW")
		self.columnVar.set(self.columnsHeader[0])

		## ROW 2 ##

		currentRow += 1

		# dictionary input entry label
		dictionaryLabel = tk.Label(self, anchor="w", fg="black", text="Dictionary to use: ")
		dictionaryLabel.grid(column=0, row=currentRow, sticky="EW")

		# dictionary input entry
		self.dictEntry = tk.Entry(self, textvariable = self.dictVarConform, width = entryWidth)
		self.dictEntry.grid(column=1, row=currentRow, sticky="EW")

		# dictionary browse button
		browseDictionaryButton = tk.Button(self, text="Browse...", command=self.OnBrowseDictButtonClick)
		browseDictionaryButton.grid(column=2, row=currentRow)

		## radio button ##
		currentRow += 1

		fullNameButton = tk.Radiobutton(self, text="Use Full Name", variable=self.radioVar, \
			value=functions.FULL_NAME_FLAG(), command=self.ToggleFullName)
		abbrNameButton = tk.Radiobutton(self, text="Use Abbreviated Name", variable=self.radioVar, \
			value=functions.ABBR_NAME_FLAG(), command=self.ToggleAbbrName)

		fullNameButton.grid(column=0, row=currentRow)
		abbrNameButton.grid(column=1, row=currentRow)

		## ROW 3 ##
		currentRow += 1

		# CONFORM button
		conformButton = tk.Button(self, text="Start Conforming", command=self.OnConformClick)
		conformButton.grid(column=0, row=currentRow)

		## ROW 4 ##

		currentRow += 1

		# Instruction label
		instructionLabel = tk.Message(self, anchor="center", justify="center", width=550, fg="black", text="Before merging, update the non_match file with the desired mapping. " +
			"Save and then close this file. " + 
			"If a conform step happens before this, the fields below will be automatically populated with the corresponding files.")
		instructionLabel.grid(column=0, row=currentRow, columnspan=2)

		## ROW 5 ## 

		currentRow += 1

		# non_match file input label
		non_matchedLabel = tk.Label(self, anchor="w", fg="black", text="Non_match file to merge: ")
		non_matchedLabel.grid(column=0, row=currentRow, sticky="EW")

		# non_match file input entry
		self.non_matchedEntry = tk.Entry(self, textvariable = self.inputVarMerge, width = entryWidth)
		self.non_matchedEntry.grid(column=1, row=currentRow, sticky="EW")

		# non_match file browse button
		nonmatch_fileButton = tk.Button(self, text="Browse...", command=self.OnBrowseInputButtonClickForMerge)
		nonmatch_fileButton.grid(column=2, row=currentRow)

		## ROW 6 ## 

		currentRow += 1

		# dictionary input entry label
		dictionaryLabel2 = tk.Label(self, anchor="w", fg="black", text="Dictionary to use: ")
		dictionaryLabel2.grid(column=0, row=currentRow, sticky="EW")

		# dictionary input entry
		self.dictEntry2 = tk.Entry(self, textvariable = self.dictVarMerge, width = entryWidth)
		self.dictEntry2.grid(column=1, row=currentRow, sticky="EW")

		# dictionary browse button
		dictionaryBrowseButtonForMerge = tk.Button(self, text="Browse...", command=self.OnBrowseDictButtonClickForMerge)
		dictionaryBrowseButtonForMerge.grid(column=2, row=currentRow)

		## ROW 7 ## 

		currentRow += 1

		# MERGE button
		mergeButton = tk.Button(self, text="Start Merging", command=self.OnMergeClick)
		mergeButton.grid(column=0, row=currentRow)

		## ROW 8 ## 

		currentRow += 1

		quitButton = tk.Button(self, text="Close Application", command=self.OnQuitClick)
		quitButton.grid(column=0, row=currentRow)

		self.grid_columnconfigure(0, weight = 2)
 	
 	# handles browse button to select file to conform
	def OnBrowseInputButtonClick(self):
		root = tk.Tk()
		root.withdraw()
		user_input = filedialog.askopenfilename(title = "Choose file to conform")
		if user_input:
			if not functions.readInputFile(user_input, functions.CONFORM_FLAG()):
				# messagebox.showerror("Read File", "Unable to read file. Please try again")
				self.inputVarConform.set("")
				return

			self.inputVarConform.set(user_input)

			# populate the options for the options menu
			self.columnsHeader = functions.getHeader()
			self.columnVar.set(self.columnsHeader[0])
			self.columnMenu["menu"].delete(0, "end")
			for item in self.columnsHeader:
				self.columnMenu["menu"].add_command(label=item, command=tk._setit(self.columnVar, item))
		else:
			messagebox.showwarning("Warning!", "No file selected")

	# handles browse button to select dictionary
	def OnBrowseDictButtonClick(self):
		root = tk.Tk()
		root.withdraw()
		dictionary_input = filedialog.askopenfilename(title = "Choose dictionary")
		if dictionary_input:
			if not functions.readDictionaryFile(dictionary_input, functions.CONFORM_FLAG()):				
				self.dictVarConform.set("")
				return

			self.dictVarConform.set(dictionary_input)
		else:
			messagebox.showwarning("Warning!", "No file selected")

	# toggle to use Full Name
	def ToggleFullName(self):
		self.radioVar.set(functions.FULL_NAME_FLAG())

	def ToggleAbbrName(self):
		self.radioVar.set(functions.ABBR_NAME_FLAG())

	# toggle to use Abbreviated Name

	# start the conform functionality
	def OnConformClick(self):

		# user has selected valid files as arguments
		if self.dictVarConform.get() and self.inputVarConform.get():
			functions.conform(self.columnVar.get(), self.radioVar.get())			
			self.inputVarMerge.set(functions.getNonMatchFilePath())
			self.dictVarMerge.set(self.dictVarConform.get())

			# clear the current conform field to force reselect of input file
			self.inputVarConform.set("")
			self.columnVar.set("(empty)")
			self.columnMenu["menu"].delete(0, "end")
			self.columnMenu['menu'].add_command(label="(empty)", command=tk._setit(self.columnVar, ["(empty)"]))
			self.dictVarConform.set("")

			messagebox.showinfo("Success!", "CONFORM PROCESS COMPLETED! Go to terminal window for more info...")

		# one of more of the files are not valid
		else:			
			if not self.inputVarMerge:
				messagebox.showerror("Invalid arguments!", "Please make sure to select the necessary files!")

	# handles the input entry for non_matched file for merge
	def OnBrowseInputButtonClickForMerge(self):
		root = tk.Tk()
		root.withdraw()
		user_input = filedialog.askopenfilename(title = "Choose file to update")

		# user has selected some file
		if user_input:
			if user_input.endswith(".xlsx") or user_input.endswith(".xls") or user_input.endswith(".csv"):
				self.inputVarMerge.set(user_input)
			else:
				messagebox.showerror("Error", "Unsupported file extension. Please choose .xlsx, .xls or .csv files!")

		# no file is selected
		else:			
			if not self.inputVarMerge:
				messagebox.showwarning("Warning!", "No file selected")

	# handles the input entry for dictionary to merge
	def OnBrowseDictButtonClickForMerge(self):
		root = tk.Tk()
		root.withdraw()
		dictionary_input = filedialog.askopenfilename(title = "Choose dictionary")

		# user has selected some file for dictionary
		if dictionary_input:
			if dictionary_input.endswith(".xlsx") or dictionary_input.endswith(".xls") \
				or dictionary_input.endswith(".csv"):
				self.dictVarMerge.set(dictionary_input)
			else:
				messagebox.showerror("Error", "Unsupported file extension. Please choose .xlsx, .xls or .csv files!")

		# no file is selected
		else:
			messagebox.showwarning("Warning!", "No file selected")

	# starts the merging functionality
	def OnMergeClick(self):

		# only perform merge if both non_match file and dictionary file
		# have been specified. First read in the input file, and if 
		# successful, read in the dictionary file. On success, perform
		# merging.
		if self.inputVarMerge.get() and self.dictVarMerge.get():
			if not functions.readInputFile(self.inputVarMerge.get(), functions.MERGE_FLAG()):
				messagebox.showerror("Read File", "Unable to read file. Please try again")
				self.inputVarMerge.set("")
			else:
				if not functions.readDictionaryFile(self.dictVarMerge.get(), functions.MERGE_FLAG()):	
					messagebox.showerror("Read File", "Unable to read file. Please try again")
					self.dictVarMerge.set("")
					return
				else:			
					functions.merge()
					messagebox.showinfo("Success!", "MERGE PROCESS COMPLETED! Go to terminal window for more info...")
		else:			
			messagebox.showerror("Error!", "Unable to perform operation because of invalid arguments! Please try again!")

	# delete all intermediary files then quit the application
	def OnQuitClick(self):
		functions.clean_temp_file()
		quit()

# main
if __name__ == "__main__":

	# delete all intermediary files then quit the application
	def on_closing():
		functions.clean_temp_file()
		quit()

	app = App(None)
	app.master.title("Name Conforming Application")
	app.master.minsize(700, 300)
	app.master.protocol("WM_DELETE_WINDOW", on_closing)
	app.mainloop()