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
# Name    : OverviewFilterSet.py
# Date    : 01/01/2020
# Revision: 1.0
# Purpose : OverviewFilterSet
#
# Usage   : NetAn RNO MRR
#

from Spotfire.Dxp.Application.Visuals import *
from Spotfire.Dxp.Data import *

myVis = myVis.As[Visualization]()

filter_option = Document.Properties["OverviewFilter"]
operator_value = Document.Properties["overviewOperatorFilter"]
filter_target = str(Document.Properties["overviewFilterValue"])

Document.Properties["overviewFilterValue"] = 0.00

