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
# Name    : RefreshTable.py
# Date    : 01/01/2020
# Revision: 1.0
# Purpose : RefreshTable
#
# Usage   : NetAn RNO MRR
#

import clr
from System.Collections.Generic import List, Dictionary
from Spotfire.Dxp.Data import DataTable
from Spotfire.Dxp.Application.Scripting import ScriptDefinition
from Spotfire.Dxp.Framework.ApplicationModel import NotificationService
from Spotfire.Dxp.Data import CalculatedColumn
from Spotfire.Dxp.Data import *
import System
from System.IO import StringReader, StreamReader, StreamWriter, MemoryStream, SeekOrigin
from System import Array
from Spotfire.Dxp.Data import Import
from Spotfire.Dxp.Data.Import import TextFileDataSource, TextDataReaderSettings
from Spotfire.Dxp.Application.Visuals import *
from Spotfire.Dxp.Framework.ApplicationModel import *
import time

def create_cursor(data_table):
    """Create cursors for a given table, these are used to loop through columns"""
    curs_list = []
    col_list = []
    for curr_col in data_table.Columns:
        curs_list.append(DataValueCursor.CreateFormatted(data_table.Columns[curr_col.Name]))
        col_list.append(data_table.Columns[curr_col.Name].ToString())
    cusr_dict = dict(zip(col_list, curs_list))
    return cusr_dict


def remove_columns(table_cols):
    """ takes a list of columns and deletes them."""
    cols_to_delete = List[DataColumn]()

    calc_col_list = [
        'Difference',
        'File ID',
        'Group ID'
    ]

    try:
        for column in table_cols:
            for check_name in calc_col_list:
                if check_name in column.Name:
                    cols_to_delete.Add(column)

        table_cols.Remove(cols_to_delete)
    except Exception as e:
        print e.message


def create_loaded_recordings_table(data,destination_table):
    """ generates a datatable given a dataset and a destination table"""
    stream = MemoryStream()
    writer = StreamWriter(stream)
    writer.Write(data)
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

    data_source = TextFileDataSource(stream, readerSettings)

    if Document.Data.Tables.Contains(destination_table):
        dataTable = Document.Data.Tables[destination_table]
        dataTable.ReplaceData(data_source)
    else:
        # Create Table if not already present
        dataTable = Document.Data.Tables.Add(destination_table, data_source)


def execute():
    try:
        ps.CurrentProgress.ExecuteSubtask("Refresh all data tables");

        # Refresh all tables. This will be done in background threads
        for t in tables:
            t.Refresh()
        


        # Wait until all tables ahev been refreshed
        working = True

        while working:
            time.sleep(2)

            # Check if user has canceled
            ps.CurrentProgress.CheckCancel()

            # Check if any data table is still refreshing
            working = False
            for t in tables:
                if t.Refreshing:
                    working = True
                    break

    except: # user canceled
        pass


Document.Properties["errorMessage"] = ""
Document.Properties["ActionMessage"] = ""
Document.Properties["removeGroupMessage"] = ""

tables = List[DataTable]()

grouped_recording = False

page = Application.Document.ActivePageReference
if page.Title == 'MRR Recordings - Single Recording':
     grouped_recording = False
else:
    grouped_recording = True

# defines what columns, tables etc. to use for single/grouped recordings
if not grouped_recording:
    tables.Add(Document.Data.Tables["combinedRecordingData"]) 
    loaded_data_table = 'combinedRecordingData'
    grouped_recording = False
    destination_table = 'Loaded Single Recordings'
    source_table_single_recording = source_table_single_recording.As[Visualization]()
    source_table = source_table_single_recording.Data.DataTableReference
    marking_name = "LandingPageMarking"
    recording_columns = ['FILE_ID','BSC_Node_Name','Recording_Start_time','Recording_Stop_time','RID','UTC_DATETIME_ID','TTIME','DTMFILTER_String','CONTYPE_String','VAMOSFILTER_String']

else:
    tables.Add(Document.Data.Tables["combinedRecordingDataGrouping"])
    loaded_data_table = 'combinedRecordingDataGrouping'
    destination_table = "Grouped_Recordings"
    source_table_name = "All_Grouped_Recordings"
    marking_name = "markedGrouping"
    recording_columns = ['GroupingName','FILE_ID','BSC_Node_Name','Recording_Start_time','Recording_Stop_time','UTC_DATETIME_ID']
    source_table = Document.Data.Tables[source_table_name]


# remove all rows previously loaded to ensure nothing shows up
Document.Data.Tables[loaded_data_table].RemoveRows(RowSelection(IndexSet(Document.Data.Tables[loaded_data_table].RowCount,True))) 

markings = Document.Data.Markings[marking_name].GetSelection(source_table)

# create cursor for table to loop through column values
column_list = create_cursor(source_table)

# get the column header as the first row
data_row = ''
text_strings = []

for column in recording_columns:
    text_strings.append(column)
data_row = '\t'.join(text_strings)
data_row += "\r\n"

column_header_string = data_row

# if there are markings selected then loop through the source table (single or grouped) and create strings for each row
if markings.IncludedRowCount != 0:
    for row in source_table.GetRows(markings.AsIndexSet(), Array[DataValueCursor](column_list.values())):
        row_string = []
        for column in recording_columns:
            row_string.append(column_list[column].CurrentValue)

        new_row_of_data = '\t'.join(row_string)
        data_row += new_row_of_data + "\r\n"
    if data_row != column_header_string: # i.e. if there have been more rows added to data_row, then create a table to show loaded recordings
        create_loaded_recordings_table(data_row,destination_table)

# in the case of a grouped recording need to highlight all rows in order for the join to work on the table
if grouped_recording:
    table = Document.Data.Tables["Grouped_Recordings"]
    row_count = table.RowCount
    rows_to_mark = IndexSet(row_count, True)
    Document.Data.Markings[marking_name].SetSelection(RowSelection(rows_to_mark), table)

# remove difference columns from comparison report
difference_columns = Document.Data.Tables[loaded_data_table].Columns
remove_columns(difference_columns)

#load the refrsh tables
ps = Application.GetService[ProgressService]()

if markings.IsEmpty:
    Document.Properties["loadError"] = "* Select item(s) from table to load"
else:
    ps.ExecuteWithProgress("Loading...", "Refreshing Data Table...", execute)
    Document.Properties["loadError"] = ""
