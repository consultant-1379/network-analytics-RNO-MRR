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
# Name    : RefreshRecordings.py
# Date    : 01/01/2020
# Revision: 1.0
# Purpose : RefreshRecordings
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
        Document.Properties["loadError"] = "User Cancelled Refresh.."


Document.Properties["errorMessage"] = ""
Document.Properties["ActionMessage"] = ""
Document.Properties["removeGroupMessage"] = ""
Document.Properties["loadError"] = ""

tables = List[DataTable]()


tables.Add(Document.Data.Tables["IL_DC_E_NETOP_MRR_ADM_RAW"])

#load the refrsh tables
ps = Application.GetService[ProgressService]()
ps.ExecuteWithProgress("Loading...", "Refreshing Data Table...", execute)



