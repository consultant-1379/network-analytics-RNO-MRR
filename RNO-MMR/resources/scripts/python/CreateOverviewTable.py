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
# Name    : CreateOverviewTable.py
# Date    : 27/11/2019
# Revision: 1.0
# Purpose : Create Overview Table
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
import time

def removeColumns(tableCols):
    """ takes a list of columns and deletes them."""
    columnsToDelete = List[DataColumn]()

    calcColumnNameList = [
            'Difference',
            'File ID',
            'Group ID'
        ]

    try:
        for column in tableCols:
            for checkName in calcColumnNameList:
                if checkName in column.Name:
                    columnsToDelete.Add(column)

        tableCols.Remove(columnsToDelete)
    except Exception as e:
        print e.message


def get_visual_ref(page_name, visual_name):

    for page in Document.Pages:
        if page.Title == page_name:
            for visual in page.Visuals:
                if visual.Title == visual_name:
                    visual_ref = visual.As[Visualization]()

    return visual_ref


Document.Properties["errorMessage"] =''
Document.Properties["removeGroupMessage"] = ""

# remove comparison chart axis as this casuses slow down
comparisonTable = get_visual_ref('MRR Comparison Report','Comparison Report Table')
comparisonTable.MeasureAxis.Expression = ''

OverviewTable = Document.Data.Tables["OverviewTable"]

columnsToDelete = List[DataColumn]()

dataResolution = Document.Properties["DataResolution"]

Document.Properties["overviewFilterValue"] = 0.0

for column in OverviewTable.Columns:
    if column.Name != 'IsDuplicate':
        columnsToDelete.Add(column)

OverviewTable.Columns.Remove(columnsToDelete)

sourceDataTableName = ""

page = Application.Document.ActivePageReference

if page.Title == 'MRR Recordings - Single Recording':
     sourceDataTableName = "combinedRecordingData"
     Document.Properties["currentLoadedTable"] = 'single'
     Document.Properties['histogramType'] = 'single_histogram'
else:
    sourceDataTableName = "combinedRecordingDataGrouping"
    Document.Properties["currentLoadedTable"] = 'group'
    Document.Properties['histogramType'] = 'group_histogram'


# remove out comparison columns from required source table
dataColumns = Document.Data.Tables[sourceDataTableName].Columns
removeColumns(dataColumns)

# copy data from source table to overview table (used as source for comparison)
sourceDataTable = Document.Data.Tables[sourceDataTableName]

dataSource=DataTableDataSource(sourceDataTable)

OverviewTable.ReplaceData(dataSource)

Document.Properties["KPIList"] = ""
Document.Properties["KPIListOT"] = ""

myVis = myVis.As[Visualization]()
dataTable = myVis.Data.DataTableReference

myVis.TableColumns.Clear()
myVis.TableColumns.Add(dataTable.Columns.Item["BSC_Node_Name"])
myVis.TableColumns.Add(dataTable.Columns.Item["CellName"])
myVis.TableColumns.Add(dataTable.Columns.Item["CELL_BAND"])
if Document.Properties["DataResolution"] != "Cell":
    myVis.TableColumns.Add(dataTable.Columns.Item["ChannelGroup"])
myVis.TableColumns.Add(dataTable.Columns.Item["SubCell_String"])


if dataResolution == "Cell":
    myVis.TableColumns.Remove(dataTable.Columns.Item["ChannelGroup"])

# set up filter vals
filter_string = "[IsDuplicate] != TRUE"
myVis.Data.WhereClauseExpression = filter_string

Document.Properties["OverviewFilter"] = ''
Document.Properties["overviewOperatorFilter"] = '='
Document.Properties["overviewFilterValue"] = 0.0

for page in Document.Pages:
    if (page.Title == 'MRR Overview Report'):
        Document.ActivePageReference=page

