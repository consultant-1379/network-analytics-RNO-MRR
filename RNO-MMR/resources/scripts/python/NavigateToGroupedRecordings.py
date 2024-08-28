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
# Name    : NavigateToGroupedRecordings.py
# Date    : 01/01/2020
# Revision: 1.0
# Purpose : NavigateToGroupedRecordings
#
# Usage   : NetAn RNO MRR
#
for page in Document.Pages:
    if (page.Title == 'MRR Recordings - Grouped Recordings'):
        Document.ActivePageReference=page

Document.Properties["errorMessage"] = ""
Document.Properties["ActionMessage"] = ""
Document.Properties["removeGroupMessage"] = ""
Document.Properties["loadError"] = ""