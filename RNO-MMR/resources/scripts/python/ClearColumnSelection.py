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
# Name    : ClearColumnSelection.py
# Date    : 01/01/2020
# Revision: 1.0
# Purpose : ClearColumnSelection
#
# Usage   : NetAn RNO MRR
#

from Spotfire.Dxp.Application.Visuals import *

Document.Properties["KPIList"] = ""

myVis = myVis.As[Visualization]()
dataTable = myVis.Data.DataTableReference

myVis.TableColumns.Clear()
myVis.TableColumns.Add(dataTable.Columns.Item["BSC_Node_Name"])
myVis.TableColumns.Add(dataTable.Columns.Item["CellName"])
myVis.TableColumns.Add(dataTable.Columns.Item["CELL_BAND"])
myVis.TableColumns.Add(dataTable.Columns.Item["SubCell_String"])
if Document.Properties["DataResolution"] != "Cell":
    myVis.TableColumns.Add(dataTable.Columns.Item["ChannelGroup"])


