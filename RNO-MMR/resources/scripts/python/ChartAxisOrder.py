# ********************************************************************
# Ericsson Inc.                                                 SCRIPT
# ********************************************************************
#
#
# (c) Ericsson Inc. 2020 - All rights reserved.
#
# The copyright to the computer program(s) herein is the property
# of Ericsson Inc. The programs may be used and/or copied only with
# the written permission from Ericsson Inc. or in accordance with the
# terms and conditions stipulated in the agreement/contract under
# which the program(s) have been supplied.
#
# ********************************************************************
# Name    : ChartAxisOrder.py
# Date    : 01/01/2020
# Revision: 1.0
# Purpose : ChartAxisOrder
#
# Usage   : NetAn RNO MRR
#
from Spotfire.Dxp.Application.Visuals import *
from System.Drawing import Color
from Spotfire.Dxp.Data import *
from System.Collections.Generic import List

order = Document.Properties['Order']
print order
recording = Document.Properties['FileOrGroupID']
print recording

sourceDataTableName = "OverviewTableUnpivot"
if Document.Properties["FileCount"] == 2 and Document.Properties["GroupingCount"] == 0:
    columnName = "FILE_ID"
elif Document.Properties["GroupingCount"] == 2 and Document.Properties["FileCount"] == 0:
    columnName = "GroupingName"

sourceDataTable = Document.Data.Tables[sourceDataTableName]

cursor = DataValueCursor.CreateFormatted(sourceDataTable.Columns[columnName])
#list object to store retrieved values
valData = List [str]()

#iterate through table column rows to retrieve the values
for row in sourceDataTable.GetRows(cursor):
    value = cursor.CurrentValue
    if value <> str.Empty:
        valData.Add(value)

#print only unique values
valData = List [str](set(valData))
print valData

for  vis in Document.ActivePageReference.Visuals:
	if  vis.TypeId == VisualTypeIdentifiers.CombinationChart:
		#vc = vis.As[BarChart]()
		#vc.SortedBars = 1
		#Define sorting typesmyCategoryKey0=CategoryKey() #Same as 'None'
		vc = vis.As[CombinationChart]()
		print vc.YAxis.Expression
		axis = vc.YAxis.Expression
		axis = axis.split('[')[-1]
		print axis
		axis = axis.split(']')[0]
		print axis
		myCategoryKey1=CategoryKey(axis, valData[0])
		myCategoryKey2=CategoryKey(axis, valData[1])
		print myCategoryKey1
		print myCategoryKey2
		#Perform sorting
		if order == 'Ascending' and recording == "Recording 1":
			vc.SortBy=myCategoryKey1
			#vc.Bars.ReverseSegmentOrder = False
			vc.XAxis.Reversed = False
		elif order == 'Ascending' and recording == "Recording 2":
			vc.SortBy=myCategoryKey2
			vc.XAxis.Reversed = False
		elif order == 'Descending' and recording == "Recording 1":
			vc.SortBy=myCategoryKey1
			vc.XAxis.Reversed = True
			#vc.Bars.ReverseSegmentOrder = True
		elif order == 'Descending' and recording == "Recording 2":
			vc.SortBy=myCategoryKey2
			vc.XAxis.Reversed = True





