#!/usr/bin/python

import xlrd
import xlwt
import xlutils.copy

# Variable definitions
path_to_file			= "Daten.xlsx"
ind_transactions		= 0
ind_store			= 1
col_store			= 3
col_salesperson			= 8
dict_stores_salespersons	= {}
dict_stores_num_salesperson	= {}
row_start			= 1
row_current			= row_start


###################################### Main ######################################
sheets				= xlrd.open_workbook(path_to_file)
sheet_transactions		= sheets.sheet_by_index(ind_transactions)
sheet_stores			= sheets.sheet_by_index(ind_store)

# Initialize dictionary with empty lists
for store in sheet_stores.col_values(0, 1):
	dict_stores_salespersons[store] = []

# Fill dictionary with salesperson for each store
for store in sheet_transactions.col_values(col_store, row_start):
	salesperson = sheet_transactions.cell_value(row_current, col_salesperson)
	dict_stores_salespersons[store].append(salesperson)
	row_current += 1

# Create dict with number of unique salespersons for each store
for store in dict_stores_salespersons:
	num_salesperson = len(set(dict_stores_salespersons[store]))
	dict_stores_num_salesperson[store] = num_salesperson
	print "number of unique salespersons in store %.0f: %d" % (store, dict_stores_num_salesperson[store])


##################################### Output #####################################
sheets_copied = xlutils.copy.copy(sheets)
write_sheet_stores = sheets_copied.get_sheet(ind_store)

# Create new entries in Store_File
write_sheet_stores.write(0, sheet_stores.ncols, 'num_salesperson')

for row in range(1, sheet_stores.nrows):
	store = sheet_stores.cell_value(row, 0)
	write_sheet_stores.write(row, sheet_stores.ncols, dict_stores_num_salesperson[store])

# Save to path_to_file
sheets_copied.save(path_to_file)
