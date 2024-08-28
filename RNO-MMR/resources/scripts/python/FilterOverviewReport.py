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
# Name    : FilterOverviewReport.py
# Date    : 01/01/2020
# Revision: 1.0
# Purpose : FilterOverviewReport
#
# Usage   : NetAn RNO MRR
#

from Spotfire.Dxp.Application.Visuals import *
from Spotfire.Dxp.Data import *

myVis = myVis.As[Visualization]()

where_clause = ''

filter_option = Document.Properties["OverviewFilter"]
operator_value = Document.Properties["overviewOperatorFilter"]
filter_target = str(Document.Properties["overviewFilterValue"])

if filter_option != None:
    where_clause = 'and ' + 'ROUND(['+ filter_option + '],2)' + operator_value + filter_target
else:
    Document.Properties["overviewFilterValue"] = 0.0

filter_string = "[IsDuplicate] != TRUE {0}".format(where_clause)

print Document.Properties["overviewFilterValue"]
myVis.Data.WhereClauseExpression = filter_string
print filter_string
print where_clause
print type(filter_target)