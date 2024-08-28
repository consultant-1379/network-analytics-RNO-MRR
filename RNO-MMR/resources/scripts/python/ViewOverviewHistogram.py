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
            SUM([MeasureValue]) OVER ([Measure],[Measurement_Type]) / Sum([MeasureValue]) over ([Measurement_Type]) * 100 / [TAVAL_Diff]
            WHEN ([Measurement_Type]="PLOSSDL") or ([Measurement_Type]="PLOSSUL") THEN
            SUM([MeasureValue]) OVER ([Measure],[Measurement_Type]) / Sum([MeasureValue]) over ([Measurement_Type]) * 100 / [PL_Diff]
            END"""),

    ('Cell_DiffCalc',
        """case  when [Measurement_Type]="TAVAL" THEN
            SUM([MeasureValue]) OVER ([Measure],[Measurement_Type],[CellName]) / Sum([MeasureValue]) over ([Measurement_Type],[CellName]) * 100 / [TAVAL_Diff]
            WHEN ([Measurement_Type]="PLOSSDL") or ([Measurement_Type]="PLOSSUL") THEN
            SUM([MeasureValue]) OVER ([Measure],[Measurement_Type],[CellName]) / Sum([MeasureValue]) over ([Measurement_Type],[CellName]) * 100 / [PL_Diff]
            END"""),

    ('ChGr_DiffCalc',
        """case  when [Measurement_Type]="TAVAL" THEN
            SUM([MeasureValue]) OVER ([Measure],[Measurement_Type],[CellName],[ChannelGroup]) / Sum([MeasureValue]) over ([Measurement_Type],[CellName],[ChannelGroup]) * 100 / [TAVAL_Diff]
            WHEN ([Measurement_Type]="PLOSSDL") or ([Measurement_Type]="PLOSSUL") THEN
            SUM([MeasureValue]) OVER ([Measure],[Measurement_Type],[CellName],[ChannelGroup]) / Sum([MeasureValue]) over ([Measurement_Type],[CellName],[ChannelGroup]) * 100 / [PL_Diff]
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

    ('Cell - % of Samples',
        """case  WHEN [Measurement_Type]="TAVAL" THEN 
        IF([TAVAL_Diff]=0,
        Sum([MeasureValue]) OVER ([Measure],[Measurement_Type],[CellName]) / Sum([MeasureValue]) over ([Measurement_Type],[CellName]) * 100,
        [Cell_DiffCalc]) 
        WHEN ([Measurement_Type]="PLOSSUL") OR ([Measurement_Type]="PLOSSDL") THEN 
        IF([PL_Diff]=0,
        Sum([MeasureValue]) OVER ([Measure],[Measurement_Type],[CellName]) / Sum([MeasureValue]) over ([Measurement_Type],[CellName]) * 100,
        [Cell_DiffCalc])
        ELSE 
        Sum([MeasureValue]) OVER ([Measure],[Measurement_Type],[CellName]) / Sum([MeasureValue]) over ([Measurement_Type],[CellName]) * 100
        END"""),

    ('Cell Accumulated - % of Samples',
        """If([Measurement_Type]="BSPOWER",
            Sum([MeasureValue]) OVER (Intersect([Measurement_Type],[CellName],AllNext([MeasureNum]))) / Sum([MeasureValue]) over ([Measurement_Type],[CellName]) * 100,
        if([Measurement_Type]="MSPOWER",
            Sum([MeasureValue]) OVER (Intersect([Measurement_Type],[CellName],AllPrevious([MSPOWER Range]))) / Sum([MeasureValue]) over ([Measurement_Type],[CellName]) * 100,
        Sum([MeasureValue]) OVER (Intersect([Measurement_Type],[CellName],AllPrevious([MeasureNum]))) / Sum([MeasureValue]) over ([Measurement_Type],[CellName]) * 100))"""),

    ('Channel Group - % of Samples',
        """case  WHEN [Measurement_Type]="TAVAL" THEN 
        IF([TAVAL_Diff]=0,
        Sum([MeasureValue]) OVER ([Measure],[Measurement_Type],[CellName],[ChannelGroup]) / Sum([MeasureValue]) over ([Measurement_Type],[CellName],[ChannelGroup]) * 100,
        [ChGr_DiffCalc]) 
        WHEN ([Measurement_Type]="PLOSSUL") OR ([Measurement_Type]="PLOSSDL") THEN 
        IF([PL_Diff]=0,
        Sum([MeasureValue]) OVER ([Measure],[Measurement_Type],[CellName],[ChannelGroup]) / Sum([MeasureValue]) over ([Measurement_Type],[CellName],[ChannelGroup]) * 100,
        [ChGr_DiffCalc])
        ELSE 
        Sum([MeasureValue]) OVER ([Measure],[Measurement_Type],[CellName],[ChannelGroup]) / Sum([MeasureValue]) over ([Measurement_Type],[CellName],[ChannelGroup]) * 100
        END"""),

    ('Channel Group Accumulated - % of Samples',
        """If([Measurement_Type]="BSPOWER",
            Sum([MeasureValue]) OVER (Intersect([Measurement_Type],[CellName],[ChannelGroup],AllNext([MeasureNum]))) / Sum([MeasureValue]) over ([Measurement_Type],[CellName],[ChannelGroup]) * 100,
           if([Measurement_Type]="MSPOWER",
            Sum([MeasureValue]) OVER (Intersect([Measurement_Type],[CellName],[ChannelGroup],AllPrevious([MSPOWER Range]))) / Sum([MeasureValue]) over ([Measurement_Type],[CellName],[ChannelGroup]) * 100,
           Sum([MeasureValue]) OVER (Intersect([Measurement_Type],[CellName],[ChannelGroup],AllPrevious([MeasureNum]))) / Sum([MeasureValue]) over ([Measurement_Type],[CellName],[ChannelGroup]) * 100))""")

])


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
        tableSettings = DataTableSaveSettings (dataTable, False, False)
        Document.Data.SaveSettings.DataTableSettings.Add(tableSettings)

def removeColumns(tableCols):
    """ takes a list of columns and deletes them. Used to remove old versions of formulas if settings change etc."""
    columnsToDelete = List[DataColumn]()
    calcColumnNameList = ['BSC - % Of Samples','BSC Accumulated - % of Samples','Cell - % of Samples','Cell Accumulated - % of Samples','Diff_Calc','Cell_DiffCalc','ChGr_DiffCalc','Channel Group - % of Samples','Channel Group Accumulated - % of Samples']

    try:
        for column in tableCols:
            for checkName in calcColumnNameList:
                if checkName in column.Name:
                    columnsToDelete.Add(column)

        tableCols.Remove(columnsToDelete)
    except Exception as e:
        print e.message


def addColumns(tableCols, formulaDict):
    """ genereic function to add calc columns to table"""
    for key, val in formulaDict.items():
        if len(str(key)) > 0 and len(str(val)) > 0:
            tableCols.AddCalculatedColumn(str(key),str(val))


def resetAllFilteringSchemes(dataTable):
     # Loop through all filtering schemes
     for filteringScheme in Document.FilteringSchemes:
          filteringScheme[dataTable].ResetAllFilters()


overviewTable= "OverviewTable"

# update the histogram formulas
histogramColumns = Document.Data.Tables['OverviewTableUnpivot'].Columns
removeColumns(histogramColumns)
addColumns(histogramColumns,histogramColumnsDict)

cellChgrTable = 'HistCellChGrColumns'
cellNameChGroupData = getDistinctColumnValues(overviewTable)

#add all to column values

cellNameChGroupData.append("(All),(All)" + '\r\n')
createDataTable(cellNameChGroupData, cellChgrTable)

mytable = Document.Data.Tables[cellChgrTable]
mycalcol = mytable.Columns.Item['FilteredChGr']

tempVar = mycalcol.Properties.Expression

replaceVar = tempVar.replace('${ComparisonOverviewHistCellList}','${OverviewHistCellList}')
mycalcCol = mytable.Columns['FilteredChGr'].As[CalculatedColumn]()

mycalcCol.Expression = replaceVar

#reset all filters for overviewunpivot as they could be changed by the top 10 chart/comparison chart
resetAllFilteringSchemes(Document.Data.Tables['OverviewTableUnpivot'])

#intialise dropdowns
Document.Properties['overviewHistCellList'] = '(All)'
Document.Properties['overviewHistogramChannelGroupList'] = '(All)'
Document.Properties["OverViewHistList"] = 'FER UL & DL'

#navigate to overview hist. page
for page in Document.Pages:
    if (page.Title == 'MRR Overview Histogram'):
        Document.ActivePageReference=page
