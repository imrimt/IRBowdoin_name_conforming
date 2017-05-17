# IRNameConforming
## Project by Bowdoin Institutional Research & Bowdoin Enterprise Systems
### Start Date: October 2016
### Finish Date: January 2017
### Project Demo: January 2017
#### Collaborators: Son D. Ngo, Steve Papaccio
#### Main Programming Language: Python 3.5

*REQUIRED PACKAGES*
- xlrd
- xlsxwriter
- openpyxl

*FILES IN THIS FOLDER*
functions.py
interface.py
TestingProcedure.xlsx

*SUBFOLDER*
old: contains old codes and config file

*RUNNING INSTRUCTIONS*

Note: make sure that Python 3.5 and the required packages are supported.

### Initial setup:

#### For MacOS User:
1. Open the terminal window and navigate to the script folder.
2. Run the command: python3 interface.py

#### For Window User:
1. Navigate to the script folder.
2. Double-click on the script: interface.py

### Conforming (for both OS users):

1. Load in the input file
2. Choose the column to conform
3. Load in dictionary file
4. Click on "Start Conforming" button
5. Results will be shown on the terminal window

Merging (for both OS users):

0. Edit the non_match file directly with new matches
1. Load in the non_match file (produced after the conforming step)
2. Load in the dictionary to merge with
3. Click on "Start Merging" button
4. Results will be shown on the terminal window


