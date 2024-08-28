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
# Name    : NavigateBackToRecordings.py
# Date    : 27/11/2019
# Revision: 1.0
# Purpose : NavigateBackToRecordings
#
# Usage   : NetAn RNO MRR
#

Document.Properties["errorMessage"] = ''
if Document.Properties["currentLoadedTable"] == 'single':
    naviagtePage = 'MRR Recordings - Single Recording'
else:
    naviagtePage = 'MRR Recordings - Grouped Recordings'

for page in Document.Pages:
    if (page.Title == naviagtePage):
        Document.ActivePageReference=page