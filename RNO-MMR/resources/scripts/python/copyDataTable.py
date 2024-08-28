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
# Name    : copyDataTable.py
# Date    : 01/01/2020
# Revision: 1.0
# Purpose : copyDataTable
#
# Usage   : NetAn RNO MRR
#

from Spotfire.Dxp.Data import AddRowsSettings
from Spotfire.Dxp.Data.Import import DataTableDataSource

#Source Data Table



sourceDataTableName = "combinedRecordingData"
sourceDataTable = Document.Data.Tables[sourceDataTableName]

# Mark rows on the source


dataSource=DataTableDataSource(sourceDataTable)
#print dataSource

if Document.Data.Tables.Contains("msmtExport"):
    msmtExportTable = Document.Data.Tables["msmtExport"]
    msmtExportTable.ReplaceData(dataSource)
else:
    #Document.Data.Tables.Remove(MRROverviewTable)
    newTable = Document.Data.Tables.Add("msmtExport", dataSource)
    destinationDataTable=Document.Data.Tables["msmtExport"]
    rowsettings=AddRowsSettings(destinationDataTable,dataSource)
    destinationDataTable.AddRows(dataSource,rowsettings)
