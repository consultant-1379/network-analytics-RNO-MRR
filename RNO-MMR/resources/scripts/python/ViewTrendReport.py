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
# Name    : ViewTrendReport.py
# Date    : 27/11/2019
# Revision: 1.0
# Purpose : ViewTrendReport
#
# Usage   : NetAn RNO MRR
#

from Spotfire.Dxp.Application.Visuals import *
import clr
import re
clr.AddReference('System.Data')  # needs to be imported before System.Data is called
import System
from System.Data import DataSet, DataTable, XmlReadMode
from Spotfire.Dxp.Data import DataType, DataTableSaveSettings
from System.IO import StringReader, StreamReader, StreamWriter, MemoryStream, SeekOrigin
from System.Collections.Generic import List
from Spotfire.Dxp.Data import Import
from Spotfire.Dxp.Data.Import import TextFileDataSource, TextDataReaderSettings
from Spotfire.Dxp.Data.Import import DatabaseDataSource, DatabaseDataSourceSettings
from Spotfire.Dxp.Data import *
from System import Array
import operator
import collections
from collections import OrderedDict
from Spotfire.Dxp.Data import AddRowsSettings, DataColumn
from Spotfire.Dxp.Data.Import import DataTableDataSource
from Spotfire.Dxp.Application.Visuals import *



def createCursor(eTable):
    """Create cursors for a given table, these are used to loop through columns

    Arguments:
        eTable {data table} -- table object

    Returns:
        cursDict -- dictionary of cursors for the given table
    """
    cursList = []
    colList = []
    for eColumn in eTable.Columns:
        cursList.append(DataValueCursor.CreateFormatted(eTable.Columns[eColumn.Name]))
        colList.append(eTable.Columns[eColumn.Name].ToString())
    cusrDict = dict(zip(colList, cursList))
    return cusrDict


def getDistinctColumnValues(overviewTable, column):
    uniqueNames = []
    dataTable = Document.Data.Tables[overviewTable]
    tableCursors = createCursor(dataTable)
    indexSet = IndexSet(dataTable.RowCount, True)

    for row in dataTable.GetRows(indexSet, tableCursors[column]):
        if tableCursors[column].CurrentValue not in uniqueNames:
            uniqueNames.append(tableCursors[column].CurrentValue)

    return uniqueNames

def createDataTable(textData, dataTableName):
    """"Used to create the measure names table"""

    stream = MemoryStream()
    writer = StreamWriter(stream)
    writer.WriteLine(dataTableName + '\r\n')
    writer.Flush()

    for line in textData:
        writer.WriteLine(line)

    writer.Flush()
    settings = TextDataReaderSettings()
    settings.AddColumnNameRow(0)
    settings.ClearDataTypes(False)
    stream.Seek(0, SeekOrigin.Begin)
    fs = TextFileDataSource(stream, settings)

    if Document.Data.Tables.Contains(dataTableName):
        dataTable = Document.Data.Tables[dataTableName]
        dataTable.ReplaceData(fs)
    else:
        # Create Table if not already present
        print "Adding Data Table"
        dataTable = Document.Data.Tables.Add(dataTableName, fs)
        print 'Done'

def resetAllFilteringSchemes():
     # Loop through all filtering schemes
     for filteringScheme in Document.FilteringSchemes:
          # Loop through all data tables
          for dataTable in Document.Data.Tables:
               # Reset all filters
               filteringScheme[dataTable].ResetAllFilters()

Document.Properties["errorMessage"] = ''
Document.Properties["removeGroupMessage"] = ""
OverviewTable = Document.Data.Tables["OverviewTable"]

resetAllFilteringSchemes()

page = Application.Document.ActivePageReference

if page.Title == 'MRR Recordings - Single Recording':
     sourceDataTableName = "combinedRecordingData"
     Document.Properties["currentLoadedTable"] = 'single'
else:
    sourceDataTableName = "combinedRecordingDataGrouping"
    Document.Properties["currentLoadedTable"] = 'group'


sourceDataTable = Document.Data.Tables[sourceDataTableName]

dataSource=DataTableDataSource(sourceDataTable)

OverviewTable.ReplaceData(dataSource)

myVis = myVis.As[Visualization]()
dataTable = myVis.Data.DataTableReference
my_vis = my_vis.As[Visualization]()

if Document.Properties["currentLoadedTable"] == 'single':
    myVis.RowAxis.Expression = "<[FILE_ID]>"
    my_vis.XAxis.Expression = "<[Recording_Start_time] NEST [FILE_ID]>"
elif Document.Properties["currentLoadedTable"] == 'group':
    myVis.RowAxis.Expression = "<[GroupingName]>"
    my_vis.XAxis.Expression = "<[Recording_Start_time] NEST [GroupingName]>"


overviewTable= "OverviewTable"

# get list of cellnames for dropdown source table
listOfCellNames = getDistinctColumnValues(overviewTable, 'CellName')

#add all to first entry
listOfCellNames.insert(0,'(All)')

#create the cellnames table for the cell name dropdown source
createDataTable(listOfCellNames,'CellNames')


for page in Document.Pages:
    if (page.Title == 'MRR Trend Report'):
        Document.ActivePageReference=page
