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
# Name    : AddToKPISelected.py
# Date    : 01/01/2020
# Revision: 1.0
# Purpose : AddToKPISelected
#
# Usage   : NetAn RNO MRR
#

from System.Collections.Generic import List, Dictionary
from Spotfire.Dxp.Data import *
from Spotfire.Dxp.Data import RowSelection, IndexSet


for item in Document.Properties["KPIList"]:
    if item not in Document.Properties["MeasureList"] and len(Document.Properties["MeasureList"].split(';')) < 9:
        if Document.Properties["MeasureList"] != '':
		    Document.Properties["MeasureList"] += ';' + item
        else:
            Document.Properties["MeasureList"] = item


counter = 1

for item in Document.Properties["MeasureList"].split(';'):
    #print item
    if len(Document.Properties["MeasureList"].split(';')) <= 9:
        Document.Properties["Measure"+str(counter)] = item
        counter+=1


print Document.Properties["Measure1"]