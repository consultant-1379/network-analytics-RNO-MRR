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
# Name    : ChangeBarChart.py
# Date    : 01/01/2020
# Revision: 1.0
# Purpose : ChangeBarChart
#
# Usage   : NetAn RNO MRR
#

from Spotfire.Dxp.Application.Visuals import *
from System.Drawing import Color

selectedMeasure = Document.Properties['Top10DropDown']
Operator = Document.Properties['Operator']
MeasureFilterValue = Document.Properties['MeasureFilterValue']

print Operator

myVis = myVis.As[Visualization]()
myVis.ColorAxis.Coloring.Clear()
colorRule = myVis.ColorAxis.Coloring.AddCategoricalColorRule()
if "UL & DL" in selectedMeasure:
    myVis.Data.WhereClauseExpression = ""
    if selectedMeasure == "Path Loss UL & DL Average (dB)":
        myVis.Data.WhereClauseExpression = '[Measure] = "Path Loss UL Average (dB)" or [Measure] = "Path Loss DL Average (dB)"'
        myVis.YAxis.Expression = 'Avg(if([Measure]="Path Loss UL Average (dB)",[MeasureValue])) as [Path Loss UL Average (dB)], Avg(if([Measure]="Path Loss DL Average (dB)",[MeasureValue])) as [Path Loss DL Average (dB)]'
    if selectedMeasure == "RXQUAL UL & DL Average (GSM)":
        myVis.Data.WhereClauseExpression = '[Measure] = "RXQUAL UL Average (GSM)" or [Measure] = "RXQUAL DL Average (GSM)"'
        myVis.YAxis.Expression = 'Avg(if([Measure]="RXQUAL UL Average (GSM)",[MeasureValue])) as [RXQUAL UL Average (GSM)], Avg(if([Measure]="RXQUAL DL Average (GSM)",[MeasureValue])) as [RXQUAL DL Average (GSM)]'
    if selectedMeasure == "FER UL & DL Average (GSM)":
        myVis.Data.WhereClauseExpression = '[Measure] = "FER UL Average (GSM)" or [Measure] = "FER DL Average (GSM)"'
        myVis.YAxis.Expression = 'Avg(if([Measure]="FER UL Average (GSM)",[MeasureValue])) as [FER UL Average (GSM)], Avg(if([Measure]="FER DL Average (GSM)",[MeasureValue])) as [FER DL Average (GSM)]'
    if selectedMeasure == "RXLEV UL & DL Average (dBm)":
        myVis.Data.WhereClauseExpression = '[Measure] = "RXLEV UL Average (dBm)" or [Measure] = "RXLEV DL Average (dBm)"'
        myVis.YAxis.Expression = 'Avg(if([Measure]="RXLEV UL Average (dBm)",[MeasureValue])) as [RXLEV UL Average (dBm)], Avg(if([Measure]="RXLEV DL Average (dBm)",[MeasureValue])) as [RXLEV DL Average (dBm)]'
else:
    myVis.Data.WhereClauseExpression = '[Measure] = "${Top10DropDown}"'
    myVis.YAxis.Expression = 'Avg([MeasureValue]) as ['+selectedMeasure+']'
if Operator == "Less Than":
	myVis.YAxis.Range = AxisRange(MeasureFilterValue, 'Null')