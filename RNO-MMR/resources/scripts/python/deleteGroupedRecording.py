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
# Name    : DeleteGroupedRecording.py
# Date    : 01/01/2020
# Revision: 1.0
# Purpose : DeleteGroupedRecording
#
# Usage   : NetAn RNO MRR
#

from Spotfire.Dxp.Data import *
from System.Threading import Thread
import clr
from System.Collections.Generic import List, Dictionary
from Spotfire.Dxp.Data import DataTable
from Spotfire.Dxp.Application.Scripting import ScriptDefinition
from Spotfire.Dxp.Framework.ApplicationModel import NotificationService
from Spotfire.Dxp.Data import CalculatedColumn
from Spotfire.Dxp.Data import *
import System
from System import Array
from System.IO import StringReader, StreamReader, StreamWriter, MemoryStream, SeekOrigin
from Spotfire.Dxp.Data import Import
from Spotfire.Dxp.Data.Import import TextFileDataSource, TextDataReaderSettings


def createCursor(eTable):
    """Create cursors for a given table, these are used to loop through columns
    """
    cursList = []
    colList = []
    for eColumn in eTable.Columns:
        cursList.append(DataValueCursor.CreateFormatted(eTable.Columns[eColumn.Name]))
        colList.append(eTable.Columns[eColumn.Name].ToString())
    cusrDict = dict(zip(colList, cursList))
    return cusrDict


def write_to_table(textData, table):
    stream = MemoryStream()
    writer = StreamWriter(stream)
    writer.Write(textData)
    writer.Flush()
    stream.Seek(0, SeekOrigin.Begin)

    readerSettings = TextDataReaderSettings()
    readerSettings.Separator = "\t"
    readerSettings.AddColumnNameRow(0)
    readerSettings.SetDataType(0, DataType.String)
    readerSettings.SetDataType(1, DataType.String)
    readerSettings.SetDataType(2, DataType.String)
    readerSettings.SetDataType(3, DataType.String)
    readerSettings.SetDataType(4, DataType.String)

    dSource = TextFileDataSource(stream, readerSettings)

    if Document.Data.Tables.Contains(table):
        dataTable = Document.Data.Tables[table]
        dataTable.ReplaceData(dSource)
    else:
        # Create Table if not already present
        dataTable = Document.Data.Tables.Add(table, dSource)


allGroupedRecordingTable=Document.Data.Tables["All_Grouped_Recordings"]
markedRowSelection = ""
markingName = "markedGrouping"
markedRowSelection = Document.Data.Markings[markingName].GetSelection(allGroupedRecordingTable)

if markedRowSelection.IsEmpty:
    Document.Properties["removeGroupMessage"] = "* Select groups to remove."
else:
    Document.Properties["removeGroupMessage"] = ""

    #remove rows from the table that holds all saved grouped recordings
    set1 = IndexSet(markedRowSelection.AsIndexSet())

    rowSelection = RowSelection(set1)
    allGroupedRecordingTable.RemoveRows(rowSelection)

    # remove all data from combined recording data group
    combinedRecordingGroupingTable = Document.Data.Tables["combinedRecordingDataGrouping"]
    combinedRecordingGroupingTable.RemoveRows(RowSelection(IndexSet(combinedRecordingGroupingTable.RowCount,True))) 

    #remove all from loaded data table
    loadedGroupingTable = Document.Data.Tables["Grouped_Recordings"]
    loadedGroupingTable.RemoveRows(RowSelection(IndexSet(loadedGroupingTable.RowCount,True))) 

    # to enuse that the recordings are deleted recreate the saved grouping table by replacing the source with current list of groups
    DictGroupRecordings = createCursor(allGroupedRecordingTable)

    #Dictionary of cursor
    Dict = createCursor(allGroupedRecordingTable)
    textData = ""
    header = "GroupingName\tFILE_ID\tBSC_Node_Name\tRecording_Start_time\tRecording_Stop_time\tUTC_DATETIME_ID\r\n"
    textData += header

    for row in allGroupedRecordingTable.GetRows(Array[DataValueCursor](Dict.values())):

        textData += Dict["GroupingName"].CurrentValue + "\t" + Dict["FILE_ID"].CurrentValue + "\t" + Dict["BSC_Node_Name"].CurrentValue + "\t" + Dict["Recording_Start_time"].CurrentValue + "\t" + Dict["Recording_Stop_time"].CurrentValue + "\t" + Dict["UTC_DATETIME_ID"].CurrentValue+ "\r\n"

    write_to_table(textData, "All_Grouped_Recordings")
    write_to_table(header, "Grouped_Recordings") # cretae empty data source for loaded grouped recordings




