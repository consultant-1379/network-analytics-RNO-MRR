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
# Name    : createRecordingGroup.py
# Date    : 27/11/2019
# Revision: 1.0
# Purpose : createRecordingGroup
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

RecordingTableName = "IL_DC_E_NETOP_MRR_ADM_RAW"
recordingTable = Document.Data.Tables[RecordingTableName]
GroupedRecordingTableName = "All_Grouped_Recordings"

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


def checkGroupUnique(newGroupName):
    uniqueNames = []
    dataTable = Document.Data.Tables[GroupedRecordingTableName]
    groupTableCursors = createCursor(dataTable)
    indexSet = IndexSet(dataTable.RowCount, True)

    for row in dataTable.GetRows(indexSet, groupTableCursors["GroupingName"]):
        if groupTableCursors["GroupingName"].CurrentValue == newGroupName :
            return False

    return True


#Dictionary of cursor
Dict = createCursor(recordingTable)

# Retrieve the marking selection

markingName = "LandingPageMarking"
markings = Document.Data.Markings[markingName].GetSelection(recordingTable)
markedRecordings = []
rangeValue = markings.IncludedRowCount

textData = "GroupingName\tFILE_ID\tBSC_Node_Name\tRecording_Start_time\tRecording_Stop_time\tUTC_DATETIME_ID\tVAMOSFILTER_String\tDTMFILTER_String\tCONTYPE_String\r\n"


cur = 0
groupVals = []
is_valid = True
# Iterate through the data table rows to retrieve the marked rows and creale a list for each row in a list
if markings.IncludedRowCount != 0:
    for row in recordingTable.GetRows(markings.AsIndexSet(), Dict["FILE_ID"], Dict["BSC_Node_Name"], Dict["Recording_Start_time"], Dict["Recording_Stop_time"],Dict["UTC_DATETIME_ID"], Dict["VAMOSFILTER_String"], Dict["DTMFILTER_String"], Dict["CONTYPE_String"]):

        markedRecording = []
        GroupingName=Document.Properties["GroupName"]
        groupUnique = checkGroupUnique(GroupingName)

        if GroupingName != '' and GroupingName.isalnum() and groupUnique:
            textData += GroupingName + "\t" + Dict["FILE_ID"].CurrentValue + "\t" + Dict["BSC_Node_Name"].CurrentValue + "\t" + Dict["Recording_Start_time"].CurrentValue + "\t" + Dict["Recording_Stop_time"].CurrentValue + "\t" + Dict["UTC_DATETIME_ID"].CurrentValue+ "\t" +Dict["VAMOSFILTER_String"].CurrentValue +"\t" + Dict["DTMFILTER_String"].CurrentValue +"\t" + Dict["CONTYPE_String"].CurrentValue +"\r\n"
            Document.Properties["errorMessage"]= ''
        else:
            Document.Properties["errorMessage"] = "Grouping name not valid or already used."

        if cur == 0:
            groupVals.append(Dict["BSC_Node_Name"].CurrentValue)
            groupVals.append(Dict["VAMOSFILTER_String"].CurrentValue)
            groupVals.append(Dict["DTMFILTER_String"].CurrentValue)
            groupVals.append(Dict["CONTYPE_String"].CurrentValue)
        else:
            if Dict["BSC_Node_Name"].CurrentValue not in groupVals:
                is_valid = False
                Document.Properties["errorMessage"] = "BSC_Node_Name must be the same"
                break
            elif Dict["VAMOSFILTER_String"].CurrentValue not in groupVals:
                is_valid = False
                Document.Properties["errorMessage"] = "VAMOSFILTER_String must be the same"
                break
            elif Dict["DTMFILTER_String"].CurrentValue not in groupVals:
                is_valid = False
                Document.Properties["errorMessage"] = "DTMFILTER_String must be the same"
                break
            elif Dict["CONTYPE_String"].CurrentValue not in groupVals:
                is_valid = False
                Document.Properties["errorMessage"] = "CONTYPE_String must be the same"
                break

        cur+=1
else:
    Document.Properties["errorMessage"]=  "No Recording Selected"

if textData != "GroupingName\tFILE_ID\tBSC_Node_Name\tRecording_Start_time\tRecording_Stop_time\tUTC_DATETIME_ID\tVAMOSFILTER_String\tDTMFILTER_String\tCONTYPE_String\r\n" and is_valid:

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

    if Document.Data.Tables.Contains(GroupedRecordingTableName):
        settings = AddRowsSettings(Document.Data.Tables[GroupedRecordingTableName],dSource)
        Document.Data.Tables[GroupedRecordingTableName].AddRows(dSource, settings)

    else:
        newTable = Document.Data.Tables.Add(GroupedRecordingTableName, dSource)
        tableSettings = DataTableSaveSettings (newTable, False, False)
        Document.Data.SaveSettings.DataTableSettings.Add(tableSettings)
    Document.Properties["errorMessage"] = Document.Properties["GroupName"] + " created in Grouped Recordings."
    Document.Properties["GroupName"] = ''
