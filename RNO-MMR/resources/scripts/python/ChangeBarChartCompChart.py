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
# Name    : ChangeBarChartCompChart.py
# Date    : 01/01/2020
# Revision: 1.0
# Purpose : ChangeBarChartCompChart
#
# Usage   : NetAn RNO MRR
#
from Spotfire.Dxp.Application.Visuals import *
from System.Drawing import Color

selectedMeasure = Document.Properties['CompChartDropDown']
print selectedMeasure

myViz = myViz.As[Visualization]()
#myViz.ColorAxis.Coloring.Clear()
#colorRule = myViz.ColorAxis.Coloring.AddCategoricalColorRule()

colorBySecondCol = ""

if Document.Properties["currentLoadedTable"] == 'single':
    colorBySecondCol = "FILE_ID"
elif  Document.Properties["currentLoadedTable"] == 'group':
    colorBySecondCol = "GroupingName"

findString = "Differ"
if findString in selectedMeasure:
	myViz.ColorAxis.Expression = "<[Axis.Default.Names]>"
else:
	myViz.ColorAxis.Expression = "<[Axis.Default.Names] NEST [" + colorBySecondCol + "]>"

myViz.ColorAxis.Coloring.Clear()
colorRule = myViz.ColorAxis.Coloring.AddCategoricalColorRule()
#axisExpr = myViz.YAxis.Expression

if "UL & DL" in selectedMeasure:
    myViz.Data.WhereClauseExpression = ""
    if selectedMeasure == "Path Loss UL & DL  Average (dB) DifferComp":
        myViz.Data.WhereClauseExpression = '[Measure] = "Path Loss UL Average (dB) DifferComp" or [Measure] = "Path Loss DL Average (dB) DifferComp"'
        myViz.YAxis.Expression = 'Avg(if([Measure]="Path Loss UL Average (dB) DifferComp",[MeasureValue])) as [Path Loss UL Average (dB) DifferComp], Avg(if([Measure]="Path Loss DL Average (dB) DifferComp",[MeasureValue])) as [Path Loss DL Average (dB) DifferComp]'
    if selectedMeasure == "RXQUAL UL & DL Average (GSM) DifferComp":
        myViz.Data.WhereClauseExpression = '[Measure] = "RXQUAL UL Average (GSM) DifferComp" or [Measure] = "RXQUAL DL Average (GSM) DifferComp"'
        myViz.YAxis.Expression = 'Avg(if([Measure]="RXQUAL UL Average (GSM) DifferComp",[MeasureValue])) as [RXQUAL UL Average (GSM) DifferComp], Avg(if([Measure]="RXQUAL DL Average (GSM) DifferComp",[MeasureValue])) as [RXQUAL DL Average (GSM) DifferComp]'
    if selectedMeasure == "FER UL & DL Average (GSM) DifferComp":
        myViz.Data.WhereClauseExpression = '[Measure] = "FER UL Average (GSM) DifferComp" or [Measure] = "FER DL Average (GSM) DifferComp"'
        myViz.YAxis.Expression = 'Avg(if([Measure]="FER UL Average (GSM) DifferComp",[MeasureValue])) as [FER UL Average (GSM) DifferComp], Avg(if([Measure]="FER DL Average (GSM) DifferComp",[MeasureValue])) as [FER DL Average (GSM) DifferComp]'
    if selectedMeasure == "RXLEV UL & DL Average (dBm) DifferComp":
        myViz.Data.WhereClauseExpression = '[Measure] = "RXLEV UL Average (dBm) DifferComp" or [Measure] = "RXLEV DL Average (dBm) DifferComp"'
        myViz.YAxis.Expression = 'Avg(if([Measure]="RXLEV UL Average (dBm) DifferComp",[MeasureValue])) as [RXLEV UL Average (dBm) DifferComp], Avg(if([Measure]="RXLEV DL Average (dBm) DifferComp",[MeasureValue])) as [RXLEV DL Average (dBm) DifferComp]'

    if selectedMeasure == "Path Loss UL & DL Average (dB)":
        myViz.Data.WhereClauseExpression = '[Measure] = "Path Loss UL Average (dB)" or [Measure] = "Path Loss DL Average (dB)"'
        myViz.YAxis.Expression = 'Avg(if([Measure]="Path Loss UL Average (dB)",[MeasureValue])) as [Path Loss UL Average (dB)], Avg(if([Measure]="Path Loss DL Average (dB)",[MeasureValue])) as [Path Loss DL Average (dB)]'
    if selectedMeasure == "RXQUAL UL & DL Average (GSM)":
        myViz.Data.WhereClauseExpression = '[Measure] = "RXQUAL UL Average (GSM)" or [Measure] = "RXQUAL DL Average (GSM)"'
        myViz.YAxis.Expression = 'Avg(if([Measure]="RXQUAL UL Average (GSM)",[MeasureValue])) as [RXQUAL UL Average (GSM)], Avg(if([Measure]="RXQUAL DL Average (GSM)",[MeasureValue])) as [RXQUAL DL Average (GSM)]'
    if selectedMeasure == "FER UL & DL Average (GSM)":
        myViz.Data.WhereClauseExpression = '[Measure] = "FER UL Average (GSM)" or [Measure] = "FER DL Average (GSM)"'
        myViz.YAxis.Expression = 'Avg(if([Measure]="FER UL Average (GSM)",[MeasureValue])) as [FER UL Average (GSM)], Avg(if([Measure]="FER DL Average (GSM)",[MeasureValue])) as [FER DL Average (GSM)]'
    if selectedMeasure == "RXLEV UL & DL Average (dBm)":
        myViz.Data.WhereClauseExpression = '[Measure] = "RXLEV UL Average (dBm)" or [Measure] = "RXLEV DL Average (dBm)"'
        myViz.YAxis.Expression = 'Avg(if([Measure]="RXLEV UL Average (dBm)",[MeasureValue])) as [RXLEV UL Average (dBm)], Avg(if([Measure]="RXLEV DL Average (dBm)",[MeasureValue])) as [RXLEV DL Average (dBm)]'
else:
    myViz.Data.WhereClauseExpression = '[Measure] = "${CompChartDropDown}"'
    myViz.YAxis.Expression = 'Avg([MeasureValue]) as ['+selectedMeasure+']'

print myViz.YAxis.Expression
print myViz.ColorAxis.Expression
