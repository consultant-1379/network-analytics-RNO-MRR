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
# Name    : ViewOverviewHistogram.py
# Date    : 27/11/2019
# Revision: 1.0
# Purpose : ViewOverviewHistogram
#
# Usage   : NetAn RNO MRR
#

import clr
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
from collections import OrderedDict
from Spotfire.Dxp.Data import CalculatedColumn 
from Spotfire.Dxp.Application.Filters import *

histogramColumnsDict = OrderedDict([
    ('Diff_Calc',
        """case  when [Measurement_Type]="TAVAL" THEN
            SUM([MeasureValue]) OVER ([Measure],[Measurement_Type],[FILE_ID]) / Sum([MeasureValue]) over ([Measurement_Type],[FILE_ID]) * 100 / [TAVAL_Diff]
            WHEN ([Measurement_Type]="PLOSSDL") or ([Measurement_Type]="PLOSSUL") THEN
            SUM([MeasureValue]) OVER ([Measure],[Measurement_Type],[FILE_ID]) / Sum([MeasureValue]) over ([Measurement_Type],[FILE_ID]) * 100 / [PL_Diff]
            END"""),

    ('Cell_DiffCalc',
        """case  when [Measurement_Type]="TAVAL" THEN
            SUM([MeasureValue]) OVER ([Measure],[Measurement_Type],[CellName],[FILE_ID]) / Sum([MeasureValue]) over ([Measurement_Type],[CellName],[FILE_ID]) * 100 / [TAVAL_Diff]
            WHEN ([Measurement_Type]="PLOSSDL") or ([Measurement_Type]="PLOSSUL") THEN
            SUM([MeasureValue]) OVER ([Measure],[Measurement_Type],[CellName],[FILE_ID]) / Sum([MeasureValue]) over ([Measurement_Type],[CellName],[FILE_ID]) * 100 / [PL_Diff]
            END"""),

    ('ChGr_DiffCalc',
        """case  when [Measurement_Type]="TAVAL" THEN
            SUM([MeasureValue]) OVER ([Measure],[Measurement_Type],[CellName],[ChannelGroup],[FILE_ID]) / Sum([MeasureValue]) over ([Measurement_Type],[CellName],[ChannelGroup],[FILE_ID]) * 100 / [TAVAL_Diff]
            WHEN ([Measurement_Type]="PLOSSDL") or ([Measurement_Type]="PLOSSUL") THEN
            SUM([MeasureValue]) OVER ([Measure],[Measurement_Type],[CellName],[ChannelGroup],[FILE_ID]) / Sum([MeasureValue]) over ([Measurement_Type],[CellName],[ChannelGroup],[FILE_ID]) * 100 / [PL_Diff]
            END"""),

    ('Cell - % of Samples',
        """case  WHEN [Measurement_Type]="TAVAL" THEN
            IF([TAVAL_Diff]=0,
            Sum([MeasureValue]) OVER ([Measure],[Measurement_Type],[CellName],[FILE_ID]) / Sum([MeasureValue]) over ([Measurement_Type],[CellName],[FILE_ID]) * 100,
            [Cell_DiffCalc])
            WHEN ([Measurement_Type]="PLOSSUL") OR ([Measurement_Type]="PLOSSDL") THEN
            IF([PL_Diff]=0,
            Sum([MeasureValue]) OVER ([Measure],[Measurement_Type],[CellName],[FILE_ID]) / Sum([MeasureValue]) over ([Measurement_Type],[CellName],[FILE_ID]) * 100,
            [Cell_DiffCalc])
            ELSE
            Sum([MeasureValue]) OVER ([Measure],[Measurement_Type],[CellName],[FILE_ID]) / Sum([MeasureValue]) over ([Measurement_Type],[CellName],[FILE_ID]) * 100
            END"""),
    ('BSC - % Of Samples',
        """case  WHEN [Measurement_Type]="TAVAL" THEN IF([TAVAL_Diff]=0,
            Sum([MeasureValue]) OVER ([Measure],[Measurement_Type]) / Sum([MeasureValue]) over ([Measurement_Type]) * 100,
            [Diff_Calc]) 
            WHEN ([Measurement_Type]="PLOSSUL") OR ([Measurement_Type]="PLOSSDL") THEN
            IF([PL_Diff]=0,
            Sum([MeasureValue]) OVER ([Measure],[Measurement_Type]) / Sum([MeasureValue]) over ([Measurement_Type]) * 100,
            [Diff_Calc])
            WHEN ([Measurement_Type]="MSPOWER") Then
             Sum([MeasureValue]) OVER ([MSPOWER Range],[Measurement_Type]) / Sum([MeasureValue]) over ([Measurement_Type]) * 100
            ELSE
            Sum([MeasureValue]) OVER ([Measure],[Measurement_Type]) / Sum([MeasureValue]) over ([Measurement_Type]) * 100
            END"""),

   ('BSC Accumulated - % of Samples',
        """If([Measurement_Type]="BSPOWER",
            Sum([MeasureValue]) OVER (Intersect([Measurement_Type],AllNext([MeasureNum]))) / Sum([MeasureValue]) over ([Measurement_Type]) * 100,
           if([Measurement_Type]="MSPOWER",
            Sum([MeasureValue]) OVER (Intersect([Measurement_Type],AllPrevious([MSPOWER Range]))) / Sum([MeasureValue]) over ([Measurement_Type]) * 100,
           Sum([MeasureValue]) OVER (Intersect([Measurement_Type],AllPrevious([MeasureNum]))) / Sum([MeasureValue]) over ([Measurement_Type]) * 100))"""),

    ('Cell Accumulated - % of Samples',
        """If(([Measurement_Type]="BSPOWER") ,
            Sum([MeasureValue]) OVER (Intersect([Measurement_Type],[CellName],[FILE_ID],AllNext([MeasureNum]))) / Sum([MeasureValue]) over ([Measurement_Type],[CellName],[FILE_ID]) * 100,
        if([Measurement_Type]="MSPOWER",
            Sum([MeasureValue]) OVER (Intersect([Measurement_Type],[CellName],[FILE_ID],AllPrevious([MSPOWER Range]))) / Sum([MeasureValue]) over ([Measurement_Type],[CellName],[FILE_ID]) * 100,
        Sum([MeasureValue]) OVER (Intersect([Measurement_Type],[CellName],[FILE_ID],AllPrevious([MeasureNum]))) / Sum([MeasureValue]) over ([Measurement_Type],[CellName],[FILE_ID]) * 100))"""),

    ('Channel Group - % of Samples',
        """case  WHEN [Measurement_Type]="TAVAL" THEN
        IF([TAVAL_Diff]=0,
        Sum([MeasureValue]) OVER ([Measure],[Measurement_Type],[CellName],[ChannelGroup],[FILE_ID]) / Sum([MeasureValue]) over ([Measurement_Type],[CellName],[ChannelGroup],[FILE_ID]) * 100,
        [ChGr_DiffCalc])
        WHEN ([Measurement_Type]="PLOSSUL") OR ([Measurement_Type]="PLOSSDL") THEN
        IF([PL_Diff]=0,
        Sum([MeasureValue]) OVER ([Measure],[Measurement_Type],[CellName],[ChannelGroup],[FILE_ID]) / Sum([MeasureValue]) over ([Measurement_Type],[CellName],[ChannelGroup],[FILE_ID]) * 100,
        [ChGr_DiffCalc])
        ELSE
        Sum([MeasureValue]) OVER ([Measure],[Measurement_Type],[CellName],[ChannelGroup],[FILE_ID]) / Sum([MeasureValue]) over ([Measurement_Type],[CellName],[ChannelGroup],[FILE_ID]) * 100
        END"""),

    ('Channel Group Accumulated - % of Samples',
        """If(([Measurement_Type]="BSPOWER"),
        Sum([MeasureValue]) OVER (Intersect([Measurement_Type],[CellName],[ChannelGroup],[FILE_ID],AllNext([MeasureNum]))) / Sum([MeasureValue]) over ([Measurement_Type],[CellName],[ChannelGroup],[FILE_ID]) * 100,
        if([Measurement_Type]="MSPOWER",
            Sum([MeasureValue]) OVER (Intersect([Measurement_Type],[CellName],[ChannelGroup],[FILE_ID],AllPrevious([MSPOWER Range]))) / Sum([MeasureValue]) over ([Measurement_Type],[CellName],[ChannelGroup],[FILE_ID]) * 100,
        Sum([MeasureValue]) OVER (Intersect([Measurement_Type],[CellName],[ChannelGroup],[FILE_ID],AllPrevious([MeasureNum]))) / Sum([MeasureValue]) over ([Measurement_Type],[CellName],[ChannelGroup],[FILE_ID]) * 100))""")

])

def removeColumns(tableCols, formulaType):
    """ takes a list of columns and deletes them."""
    columnsToDelete = List[DataColumn]()

    if formulaType == 'Histogram':
        calcColumnNameList = [
            'BSC - % Of Samples',
            'BSC Accumulated - % of Samples',
            'Cell - % of Samples',
            'Cell Accumulated - % of Samples',
            'Diff_Calc',
            'Cell_DiffCalc',
            'ChGr_DiffCalc',
            'Channel Group - % of Samples',
            'Channel Group Accumulated - % of Samples'
        ]


    try:
        for column in tableCols:
            for checkName in calcColumnNameList:
                if checkName in column.Name:
                    columnsToDelete.Add(column)

        tableCols.Remove(columnsToDelete)
    except Exception as e:
        print e.message

def addColumns(tableCols, formulaDict, formula_replace):
    """ genereic function to add calc columns to table"""
    for key, val in formulaDict:
        if formula_replace == 'group':
            val = val.replace("[FILE_ID]", "[GroupingName]")
        if len(str(key)) > 0 and len(str(val)) > 0:
			tableCols.AddCalculatedColumn(str(key),str(val))

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


def getDistinctColumnValues(overviewTable):
    uniqueNames = []
    dataTable = Document.Data.Tables[overviewTable]
    tableCursors = createCursor(dataTable)
    indexSet = IndexSet(dataTable.RowCount, True)
    

    for row in dataTable.GetRows(indexSet, tableCursors['ChannelGroup'], tableCursors['CellName']):
        curr_row_val = tableCursors['CellName'].CurrentValue + ',' + tableCursors['ChannelGroup'].CurrentValue
        if curr_row_val not in uniqueNames:
            uniqueNames.append(curr_row_val + '\r\n')

    return uniqueNames

def createDataTable(textData,dataTableName):
    """"Used to create the measure names table"""

    stream = MemoryStream()
    writer = StreamWriter(stream)
    writer.WriteLine('CellName,Channel Group' + '\r\n')
    writer.Flush()

    for line in textData:
        writer.WriteLine(line)

    writer.Flush()
    settings = TextDataReaderSettings()
    settings.Separator = ","
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


def resetAllFilteringSchemes(dataTable):
     # Loop through all filtering schemes
     for filteringScheme in Document.FilteringSchemes:
          filteringScheme[dataTable].ResetAllFilters()

formula_replace = ""
if Document.Properties["currentLoadedTable"] == 'single':
    formula_replace = "file_id"
    Document.Properties['histogramType'] = 'single_histogram'
elif  Document.Properties["currentLoadedTable"] == 'group':
    formula_replace = "group"
    Document.Properties['histogramType'] = 'group_histogram'


overviewTable= "OverviewTable"

# update the histogram formulas
histogramColumns = Document.Data.Tables['OverviewTableUnpivot'].Columns
removeColumns(histogramColumns, 'Histogram')
addColumns(histogramColumns, histogramColumnsDict.items(), formula_replace)


cellChgrTable = 'HistCellChGrColumns'
cellNameChGroupData = getDistinctColumnValues(overviewTable)
createDataTable(cellNameChGroupData, cellChgrTable)


mytable = Document.Data.Tables[cellChgrTable]
mycalcol = mytable.Columns.Item['FilteredChGr']

tempVar = mycalcol.Properties.Expression

replaceVar = tempVar.replace('${OverviewHistCellList}','${ComparisonOverviewHistCellList}')
mycalcCol = mytable.Columns['FilteredChGr'].As[CalculatedColumn]()

mycalcCol.Expression = replaceVar

#reset all filters for overviewunpivot as they could be changed by the top 10 chart/comparison chart
resetAllFilteringSchemes(Document.Data.Tables['OverviewTableUnpivot'])

#navigate to overview hist. page
for page in Document.Pages:
    if (page.Title == 'MRR Comparison Overview Histogram'):
        Document.ActivePageReference=page

Document.Properties["ComparisonOverViewHistList"] = 'FER UL & DL'
