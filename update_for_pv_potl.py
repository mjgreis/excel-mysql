# insert_update_from_excel
# 2019-12-29
#
# Branch version 2020-01-04 - Update rs_pv_potl field
#
# Update a MySQL table using information loaded from an Excel spreadsheet
# Written to update Needham residence table with information about solar
# 
# Import necessary libraries & modules
import xlrd
import mysql.connector
#
# Open Excel workbook; read sheet into "sheet"
book = xlrd.open_workbook(r"C:\Documents\Needham\Green Needham\Projects\Solarize Plus\Solarize_assessed_addresses_pct_h.xlsx")
sheet = book.sheet_by_name("Load")
#
# Open MySQL connection 
database = mysql.connector.connect (host="greenneedham.org", user = "michael", passwd = "bre82DON", db = "needham")
cursor = database.cursor()
#
# Set a query
#query = ("SELECT rs_street_number, rs_street_name from RESIDENCE r where rs_precinct = 'E' order by rs_street_name")
#
# Execute the query
#cursor.execute(query)
#
# Display selected fields from result
#for (rs_street_number, rs_street_name) in cursor:
#  print("{} {} ".format(
#    rs_street_number, rs_street_name))
#
# Set an update query
update_query = ("UPDATE needham.RESIDENCE SET rs_pv_potl = %s WHERE rs_id like %s")
#
# Loop through the rows of the spreadsheet
#  Read select values from each row into query parms (st_num, st_name)
#  Execute the update query (set rs_has_PV to "Y")
#
for r in range(1, sheet.nrows):
  st_num = str(int(sheet.cell(r,0).value))
  st_name = sheet.cell(r,1).value
  pv_potl = sheet.cell(r,6).value
  rs_id = int(sheet.cell(r,9).value)

  values = (pv_potl, rs_id)
  print (f" {rs_id} at {st_num} {st_name} has solar potential = {pv_potl} ")

# cursor.execute(update_query, values)
#
# Commit changes
# database.commit()
