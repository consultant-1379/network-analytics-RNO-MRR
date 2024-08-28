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
# Name    : AddColumnsToVis.py
# Date    : 01/01/2020
# Revision: 1.0
# Purpose : AddColumnsToVis
#
# Usage   : NetAn RNO MRR
#

from System.Collections.Generic import List, Dictionary
from Spotfire.Dxp.Data import *
from Spotfire.Dxp.Data import RowSelection, IndexSet
from Spotfire.Dxp.Application.Visuals import *


myVis = myVis.As[Visualization]()
dataTable = myVis.Data.DataTableReference

columnName = []
for item in Document.Properties["KPIListOT"]:
    if item not in columnName:
        columnName.append(item)


print columnName

for item in columnName:
    if myVis.TableColumns.Contains(dataTable.Columns.Item[item]):
        print "Column already in table"
    else:
        myVis.TableColumns.Add(dataTable.Columns.Item[item])

Document.Properties["overviewFilterValue"] = 0.00
