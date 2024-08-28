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
# Name    : UpdateHTML.py
# Date    : 01/01/2020
# Revision: 1.0
# Purpose : UpdateHTML
#
# Usage   : NetAn RNO MRR
#

from Spotfire.Dxp.Application.Visuals import *
from Spotfire.Dxp.Data import DataManager 
from Spotfire.Dxp.Data import IndexSet 
from Spotfire.Dxp.Data import RowSelection 


textArea.As[HtmlTextArea]().HtmlContent += " "  #adds a space to trigger javascript
htmlContent = textArea.As[HtmlTextArea]().HtmlContent 

grouped_recordings_table = Document.Data.Tables["All_Grouped_Recordings"]
adm_raw_table = Document.Data.Tables["IL_DC_E_NETOP_MRR_ADM_RAW"]

activeVisual = Document.ActivePageReference.ActiveVisualReference

if activeVisual.Title == "Individual Recordings":
    marking=Application.GetService[DataManager]().Markings["markedGrouping"]
    selectRows = IndexSet(grouped_recordings_table.RowCount, False)
    marking.SetSelection(RowSelection(selectRows),grouped_recordings_table)
elif activeVisual.Title == "Grouped Recordings ":
    marking=Application.GetService[DataManager]().Markings["LandingPageMarking"]
    selectRows = IndexSet(adm_raw_table.RowCount, False)
    marking.SetSelection(RowSelection(selectRows),adm_raw_table)
