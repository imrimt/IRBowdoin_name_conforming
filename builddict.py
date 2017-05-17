# ******************************************************************************
# Author: Son D. Ngo
# Filename: functions.py
# Contributor: Steve Papaccio
# Updated: 12/14/16

# Description: Script that implements functions that will be used for 
# 		interface.py, including reading files, running the conform and merging
#		steps. This file can be understood as the backend of program.
# ******************************************************************************

# Modules & Packages
import os
import re
import sys
import pandas as pd
import numpy as np
import difflib
import datetime
from tqdm import tqdm

full_name_column = []
id_column = []

def levenshtein(s1, s2):
    if len(s1) < len(s2):
        return levenshtein(s2, s1)

    # len(s1) >= len(s2)
    if len(s2) == 0:
        return len(s1)

    previous_row = range(len(s2) + 1)
    for i, c1 in enumerate(s1):
        current_row = [i + 1]
        for j, c2 in enumerate(s2):
            insertions = previous_row[j + 1] + 1 # j+1 instead of j since previous_row and current_row are one character longer
            deletions = current_row[j] + 1       # than s2
            substitutions = previous_row[j] + (c1 != c2)
            current_row.append(min(insertions, deletions, substitutions))
        previous_row = current_row
    
    return previous_row[-1]

def common_start_length(sa, sb):
    def _iter():
        for a, b in zip(sa, sb):
            if a == b:
                yield a
            else:
                return

    return len(''.join(_iter()))

# for multiprocessing (in the future)
# def process(item):
# 	result = ""
# 	mapped_full_name = []
# 	for (index, name) in full_name_column.iteritems():
# 		criteria_length = len(name)
# 		if (criteria_length * 0.9 >= common_start_length(item, name)):
# 			if (result == ""):
# 				result = "{} ({})".format(criteria_length, id_column[index])
# 			else:
# 				result += ", {} ({})".format(criteria_length, id_column[index])

# 	if "," not in result and result:
# 		result = result[:result.find("(") - 1]
# 	mapped_full_name.append(result)
# 	return mapped_full_name

def main():
	as_entered_file = sys.argv[1]
	full_name_file = sys.argv[2]

	as_entered_df = pd.read_excel(as_entered_file)
	full_name_df = pd.read_excel(full_name_file)

	as_entered_column = as_entered_df["AS ENTERED"]
	mapped_full_name = list(as_entered_df["FULL NAME"])
	full_name_column = full_name_df["FULL NAME"]
	id_column = full_name_df["ID"]

	# dict_df = pd.DataFrame(columns = ["AS ENTERED", "FULL NAME", "ABBREVIATED NAME", "DATE UPDATED"])
	dict_df = as_entered_df.copy()

	dict_df["AS ENTERED"] = as_entered_column

	# with Pool(multiprocessing.cpu_count()) as p:
	# 	mapped_full_name = p.map(process, as_entered_column)

	# go through every name in as_entered_column, and then for each of those names
	# go through the list of all possible full names, and see which ones match the best
	# criteria can range from 50% match to 90% match. Note that this match applies to
	# the length of whichever shorter string. Therefore, another criteria can be 
	# the length of whichever longer string.

	# Another note: The slower version, using difflib is much slower, but generate a lot more 
	# results. The current approach, is limited to most commong substring, instead of similarity
	# between strings. Maybe do more research on this?
	for index, item in enumerate(tqdm(as_entered_column, desc="Running Very Slowly...")):
		format_item = item.strip().lower()
		if (not pd.isnull(mapped_full_name[index])):
			continue
		result = ""
		for (name_index, name) in full_name_column.iteritems():
			format_name = name.strip().lower()
			criteria_length = min(len(format_name), len(format_item))

			# test different comparison algorithm
			# if (difflib.SequenceMatcher(None,format_name, format_item).ratio() >= 0.9):
			# if (common_start_length(format_name, format_item) >= criteria_length * 0.8):
			if (levenshtein(item, name) <= 10):
			# if (common_start_length(name, item) >= criteria_length * 0.9):				
				if (result == ""):
					result = "{} ({})".format(name, id_column[name_index])
				else:
					result += ", {} ({})".format(name, id_column[name_index])

		if "," not in result and result:
			result = result[:result.find("(") - 1]
		mapped_full_name[index] = result

	# replace the current FULL NAME column with the udpated version
	dict_df["FULL NAME"] = pd.Series(mapped_full_name)

	# write out the result
	writer = pd.ExcelWriter("{}_updated_{}.xlsx".format(as_entered_file[0:as_entered_file.find(".xlsx")], \
		datetime.datetime.today().strftime('%m-%d-%Y-%H-%M-%S')), engine='xlsxwriter')
	pd.formats.format.header_style = None
	dict_df.to_excel(writer, startcol=0, startrow=0, index=False, sheet_name="Sheet1")
	writer.save()

if __name__ == "__main__":
	main()