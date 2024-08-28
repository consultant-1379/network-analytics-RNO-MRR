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
# Name    : UpdateCompBarGraph.py
# Date    : 01/01/2020
# Revision: 1.0
# Purpose : UpdateCompBarGraph
#
# Usage   : NetAn RNO MRR
#

from Spotfire.Dxp.Application.Visuals import *
from System.Drawing import Color

selectedMeasure = Document.Properties['CompChartDropDown']
print selectedMeasure
Operator = Document.Properties['Operator2']
print Operator
MeasureFilterValue = Document.Properties['MeasureFilterValue2']
print MeasureFilterValue

myViz = myViz.As[Visualization]()
#myViz.ColorAxis.Coloring.Clear()
#colorRule = myViz.ColorAxis.Coloring.AddCategoricalColorRule()

colorBySecondCol = ""
if Document.Properties["FileCount"] == 2 and Document.Properties["GroupingCount"] == 0:
    colorBySecondCol = "FILE_ID"
elif Document.Properties["GroupingCount"] == 2 and Document.Properties["FileCount"] == 0:
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
        myViz.YAxis.Expression = 'Avg(if([Measure]="Path Loss UL Average (dB) DifferComp",[MeasureValue])) as [Path Loss UL Average (dB) DifferComp], Avg(if([Measure]="Path Loss DL Average (dB) DifferComp",[MeasureValue])) as [Path Loss DL Average (dB) DifferComp]'
    if selectedMeasure == "RXQUAL UL & DL Average (GSM) DifferComp":
        myViz.YAxis.Expression = 'Avg(if([Measure]="RXQUAL UL Average (GSM) DifferComp",[MeasureValue])) as [RXQUAL UL Average (GSM) DifferComp], Avg(if([Measure]="RXQUAL DL Average (GSM) DifferComp",[MeasureValue])) as [RXQUAL DL Average (GSM) DifferComp]'
    if selectedMeasure == "FER UL & DL Average (GSM) DifferComp":
        myViz.YAxis.Expression = 'Avg(if([Measure]="FER UL Average (GSM) DifferComp",[MeasureValue])) as [FER UL Average (GSM) DifferComp], Avg(if([Measure]="FER DL Average (GSM) DifferComp",[MeasureValue])) as [FER DL Average (GSM) DifferComp]'
    if selectedMeasure == "RXLEV UL & DL Average (dBm) DifferComp":
        myViz.YAxis.Expression = 'Avg(if([Measure]="RXLEV UL Average (dBm) DifferComp",[MeasureValue])) as [RXLEV UL Average (dBm) DifferComp], Avg(if([Measure]="RXLEV DL Average (dBm) DifferComp",[MeasureValue])) as [RXLEV DL Average (dBm) DifferComp]'
    if selectedMeasure == "Path Loss UL & DL Average (dB)":
        myViz.YAxis.Expression = 'Avg(if([Measure]="Path Loss UL Average (dB)",[MeasureValue])) as [Path Loss UL Average (dB)], Avg(if([Measure]="Path Loss DL Average (dB)",[MeasureValue])) as [Path Loss DL Average (dB)]'
    if selectedMeasure == "RXQUAL UL & DL Average (GSM)":
        myViz.YAxis.Expression = 'Avg(if([Measure]="RXQUAL UL Average (GSM)",[MeasureValue])) as [RXQUAL UL Average (GSM)], Avg(if([Measure]="RXQUAL DL Average (GSM)",[MeasureValue])) as [RXQUAL DL Average (GSM)]'
    if selectedMeasure == "FER UL & DL Average (GSM)":
        myViz.YAxis.Expression = 'Avg(if([Measure]="FER UL Average (GSM)",[MeasureValue])) as [FER UL Average (GSM)], Avg(if([Measure]="FER DL Average (GSM)",[MeasureValue])) as [FER DL Average (GSM)]'
    if selectedMeasure == "RXLEV UL & DL  Average (dBm)":
        myViz.YAxis.Expression = 'Avg(if([Measure]="RXLEV UL Average (dBm)",[MeasureValue])) as [RXLEV UL Average (dBm)], Avg(if([Measure]="RXLEV DL Average (dBm)",[MeasureValue])) as [RXLEV DL Average (dBm)]'
else:
    myViz.Data.WhereClauseExpression = '[Measure] = "${CompChartDropDown}"'
    myViz.YAxis.Expression = 'Avg([MeasureValue]) as ['+selectedMeasure+']'

if Operator == "Less Than":
	myViz.YAxis.Range = AxisRange(None, MeasureFilterValue)
elif Operator == "Greater Than":
	myViz.YAxis.Range = AxisRange(MeasureFilterValue, None)
elif Operator == "Equal":
	myViz.YAxis.Range = AxisRange(MeasureFilterValue, MeasureFilterValue)
elif Operator == "Auto":
	myViz.YAxis.Range = AxisRange(None, None)
