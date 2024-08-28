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
# Name    : MarkItems.py
# Date    : 01/01/2020
# Revision: 1.0
# Purpose : MarkItems
#
# Usage   : NetAn RNO MRR
#

from Spotfire.Dxp.Application.Visuals import VisualContent
from Spotfire.Dxp.Data import IndexSet
from Spotfire.Dxp.Data import RowSelection
from Spotfire.Dxp.Data import DataValueCursor
from System.Collections.Generic import List


# Create a cursor for the table column to get the values from.
# Add a reference to the data table in the script.
dataTable = Document.Data.Tables["All_Grouped_Recordings"]
cursor = DataValueCursor.CreateFormatted(dataTable.Columns["GroupingName"])

# Retrieve the marking selection
try:
    markings = Document.ActiveMarkingSelectionReference.GetSelection(dataTable)
    # Create a List object to store the retrieved data marking selection
    markedata = List [str]();

    # Iterate through the data table rows to retrieve the marked rows
    for row in dataTable.GetRows(markings.AsIndexSet(),cursor):
        #rowIndex = row.Index ##un-comment if you want to fetch the row index into some defined condition
        value = cursor.CurrentValue
        if value <> str.Empty:
            markedata.Add(value)

    # Get only unique values
    valData = List [str](set(markedata))

    # Store in a document property
    markedItems = ', '.join(valData)


    #Script parameters:
    selectedTag = markedItems
    ColumnName = 'GroupingName'

    #get table reference
    vc = vis.As[VisualContent]()
    dataTable = vc.Data.DataTableReference

    #set marking
    marking=vc.Data.MarkingReference

    #setup rows to select from rows to include
    rowCount=dataTable.RowCount
    rowsToInclude = IndexSet(rowCount, True)
    rowsToSelect = IndexSet(rowCount, False)

    cursor1 = DataValueCursor.CreateFormatted(dataTable.Columns[ColumnName])

    #Find records by looping through all rows
    i=0
    groups = markedItems.split(",")
    groups = [x.strip(' ') for x in groups]

    for row in dataTable.GetRows(rowsToInclude, cursor1):
        aTag = cursor1.CurrentValue
        #print aTag
        if aTag in groups: 
            rowsToSelect[i]=True
            print i

        i=i+1


    print rowsToSelect
    #Set marking
    marking.SetSelection(RowSelection(rowsToSelect), dataTable)
except Exception as e:
    print e.message


