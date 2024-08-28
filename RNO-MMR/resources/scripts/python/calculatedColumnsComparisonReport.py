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
# Name    : CalculatedColumnsOthers.py
# Date    : 01/01/2020
# Revision: 1.0
# Purpose : CalculatedColumnsOthers
#
# Usage   : NetAn RNO MRR
#
from Spotfire.Dxp.Data import *
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

convertNames = {
  'FERUL': 'FER UL',
  'FERDL': 'FER DL',
  'RXQUALUL': 'RXQUAL UL',
  'RXQUALDL': 'RXQUAL DL',
  'RXLEVUL': 'RXLEV UL',
  'RXLEVDL': 'RXLEV DL',
  'PLOSSUL': 'Path Loss UL',
  'PLOSSDL': 'Path Loss DL',
  'PLDIFF': 'Path Loss Diff.',
  'TAVAL': 'Timing Adv.',
  'BSPOWER': 'Power Red. BS',
  'MSPOWER': 'Power Lev. MS'
}

compChartColumns = {
"FER DL Average (GSM)",
"FER UL Average (GSM)",
"Path Loss Diff. Average (dB)",
"Path Loss DL Average (dB)",
"Path Loss UL Average (dB)",
"Power Lev. MS Average (dBm)",
"Power Red. BS Average (dB)",
"RXLEV DL Average (dBm)",
"RXLEV UL Average (dBm)",
"RXQUAL DL Average (GSM)",
"RXQUAL UL Average (GSM)",
"Timing Adv. Average (GSM)"
}

def createCursor(eTable):
    """Create cursors for a given table, these are used to loop through columns"""
    cursList = []
    colList = []
    for eColumn in eTable.Columns:
        cursList.append(DataValueCursor.CreateFormatted(eTable.Columns[eColumn.Name]))
        colList.append(eTable.Columns[eColumn.Name].ToString())
    cusrDict = dict(zip(colList, cursList))
    return cusrDict

# =============================================
# Diff. Formulas
# =============================================

def diffKPIcalc(table, columnName, nameString):
    diffCols = OrderedDict()
    cols = Document.Data.Tables[table.Name].Columns

    cursor = DataValueCursor.CreateFormatted(table.Columns[columnName])
    #list object to store retrieved values
    valData = List [str]()

    #iterate through table column rows to retrieve the values
    for row in table.GetRows(cursor):
        value = cursor.CurrentValue
        if value <> str.Empty:
            valData.Add(value)

    #print only unique values
    comparisonAxis = []
    valData = List [str](set(valData))
    print valData

    #print only unique values
    comparisonAxis = []

    measureTable = Document.Data.Tables["Measure Names"]
    mappingDataCur = createCursor(measureTable)
    
    for row in measureTable.GetRows(mappingDataCur["Measures"]):
		col = mappingDataCur['Measures'].CurrentValue
		colName1 = col + " " + nameString + " 1"
		colName2 = col + " " + nameString + " 2"
		diffcol = col + " Difference"
		expr1 = "case  when [" + columnName + "]='"+valData[0]+"' then [" + col + "] else 0 end"
		expr2 = "case  when [" + columnName + "]='"+valData[1]+"' then [" + col + "] else 0 end"
		expressionDiff = "[" + colName1 + "] - [" + colName2 + "]"
		exprdiff = """
                CASE  WHEN {0} IS NULL THEN
                    0
                    ELSE 
                    {0}
                    END
                """.format(expressionDiff)

		diffCols[colName1] = expr1
		diffCols[colName2] = expr2
		diffCols[diffcol] = exprdiff

		comparisonAxis.append("sum(["+colName1+"]) as ["+col+ " " +valData[0]+"]")
		comparisonAxis.append("sum(["+colName2+"]) as ["+col+ " " +valData[1]+"]")
		comparisonAxis.append("sum(["+diffcol+"]) as ["+ diffcol+"]")

    comparisonAxisValue  = ', '.join(comparisonAxis)

    return diffCols, comparisonAxisValue


def removeColumns(tableCols, formulaType):
    """ takes a list of columns and deletes them."""
    columnsToDelete = List[DataColumn]()

    if formulaType == 'Diff':
        calcColumnNameList = [
            'Difference',
            'File ID',
            'Group ID'
        ]

    elif formulaType == 'Differ':
        calcColumnNameList = [
            'DifferComp'
        ]

    try:
        for column in tableCols:
            for checkName in calcColumnNameList:
                if checkName in column.Name:
                    columnsToDelete.Add(column)

        tableCols.Remove(columnsToDelete)
    except Exception as e:
        print e.message


def addCompChartDiffColumns(tableName, columnName):
	chartDiffCols = OrderedDict()
	cols = Document.Data.Tables[tableName.Name].Columns

	cursor = DataValueCursor.CreateFormatted(tableName.Columns[columnName])
	for col in compChartColumns:
		colName = col + " DifferComp"
		colExpr = "Sum([" + col + "]) OVER (Intersect([BSC_Node_Name],[Cell Name],[CELL_BAND],[ChannelGroup],[" + columnName + "])) - Sum([" + col + "]) OVER (Intersect([BSC_Node_Name],[Cell Name],[CELL_BAND],[ChannelGroup],Next([" + columnName + "])))"
		chartDiffCols[colName] = colExpr
	return chartDiffCols

def addColumns(tableCols, formulaDict, tableName):
    """ genereic function to add calc columns to table"""
    for key, val in formulaDict:
        if tableName == 'combinedRecordingDataGrouping':
            val = val.replace("[FILE_ID]", "[GroupingName]")
        if len(str(key)) > 0 and len(str(val)) > 0:
			tableCols.AddCalculatedColumn(str(key),str(val))


def get_visual_ref(page_name, visual_name):

    for page in Document.Pages:
        if page.Title == page_name:
            for visual in page.Visuals:
                if visual.Title == visual_name:
                    visual_ref = visual.As[Visualization]()
                    found = True
    return visual_ref

def checkToplogyCount(columnName, tableName):
	flag = 0
	columnToFetch=columnName
	
	RecordingData = List [str]()
	cursor = DataValueCursor.CreateFormatted(tableName.Columns[columnToFetch])

	for row in tableName.GetRows(cursor):
		value = cursor.CurrentValue
		if value <> str.Empty:
			RecordingData.Add(value)

    #print only unique values
	RecordingData = List [str](set(RecordingData))
	
	print RecordingData
	count = 0
	for node in RecordingData: 
		count = count + 1
		print RecordingData
		print count
	if count > 1:
		flag = 1
	return flag

def resetAllFilteringSchemes():
     # Loop through all filtering schemes
     for filteringScheme in Document.FilteringSchemes:
          # Loop through all data tables
          for dataTable in Document.Data.Tables:
               # Reset all filters
               filteringScheme[dataTable].ResetAllFilters()

Document.Properties["errorMessage"] = ''
Document.Properties["removeGroupMessage"] = ""
sourceDataTableName = ""

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
OverviewTable = Document.Data.Tables["OverviewTable"]

flag = checkToplogyCount('BSC_Node_Name',sourceDataTable)

columnsToDelete = List[DataColumn]()

dataResolution = Document.Properties["DataResolution"]

for column in OverviewTable.Columns:
    if column.Name != 'IsDuplicate':
        columnsToDelete.Add(column)

OverviewTable.Columns.Remove(columnsToDelete)


if sourceDataTableName == "combinedRecordingData":
    dataColumns = Document.Data.Tables['combinedRecordingData'].Columns
    columnName = "FILE_ID"
    nameString = "File ID"
else:
    dataColumns = Document.Data.Tables['combinedRecordingDataGrouping'].Columns
    columnName = "GroupingName"
    nameString = "Group ID"

recordingsDatatable = Document.Data.Tables[sourceDataTableName]

removeColumns(dataColumns, 'Diff')
kpiFormulas, comparisonAxis= diffKPIcalc(recordingsDatatable, columnName, nameString)
addColumns(dataColumns,kpiFormulas.items(), dataSource)

removeColumns(dataColumns, 'Differ')
ChartkpiFormulas = addCompChartDiffColumns(recordingsDatatable,columnName)
addColumns(dataColumns,ChartkpiFormulas.items(), dataSource)


# make updates to chart axis etc.
Document.Properties['ComparisonAxis'] = comparisonAxis

OverviewTable.ReplaceData(dataSource)

Document.Properties["KPIList"] = ""
Document.Properties["KPIListOT"] = ""


overview_vis = get_visual_ref('MRR Overview Report','Overview Table View')

dataTable = overview_vis.Data.DataTableReference

overview_vis.TableColumns.Clear()


# create comparison chart if not exists
comparisonTable = get_visual_ref('MRR Comparison Report','Comparison Report Table')

dataResolution = Document.Properties["DataResolution"]

if dataResolution == "Cell":
    comparisonTable.RowAxis.Expression = "<[CellName] NEST [CELL_BAND]>"
else:
    comparisonTable.RowAxis.Expression = "<[CellName] NEST [CELL_BAND] NEST [ChannelGroup]>"

comparisonTable.MeasureAxis.Expression ="${ComparisonAxis}"


for page in Document.Pages:
	if flag == 1 and 'MRR Recordings' in page.Title:
		Document.ActivePageReference=page
		Document.Properties["errorMessage"] = "Please select recordings with same BSC Node Name"
		break
	elif flag == 0 and page.Title == 'MRR Comparison Report':
		Document.ActivePageReference=page
		Document.Properties["errorMessage"] = " "
	print Document.ActivePageReference.Title
