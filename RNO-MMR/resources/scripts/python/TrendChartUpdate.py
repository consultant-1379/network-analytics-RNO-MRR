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
# Name    : TrendChartUpdate.py
# Date    : 01/01/2020
# Revision: 1.0
# Purpose : TrendChartUpdate
#
# Usage   : NetAn RNO MRR
#

from Spotfire.Dxp.Application.Visuals import TablePlot
from Spotfire.Dxp.Application.Visuals import *
from System.Drawing import Color

selected_measure = Document.Properties["trendDropDown"]
selected_cell = Document.Properties["CellNameTrend"]

myVis = myVis.As[Visualization]()
tableVis = tableVis.As[Visualization]()
myVis.ColorAxis.Coloring.Clear()
colorRule = myVis.ColorAxis.Coloring.AddCategoricalColorRule()

if "UL & DL" in selected_measure:
    myVis.Data.WhereClauseExpression = "[CellName] = '"+selected_cell+"'"
    if selected_measure == "Path Loss UL & DL Average (dB)":
        myVis.YAxis.Expression = 'Avg(if([Measure]="Trend_Path Loss UL Average (dB)",[MeasureValue])) as [Path Loss UL Average (dB)], Avg(if([Measure]="Trend_Path Loss DL Average (dB)",[MeasureValue])) as [Path Loss DL Average (dB)]'
        tableVis.Data.WhereClauseExpression = '([Measure] = "Trend_Path Loss UL Average (dB)" or [Measure] = "Trend_Path Loss DL Average (dB)") and [CellName] = "'+selected_cell+'"'
    if selected_measure == "RXQUAL UL & DL Average (GSM)":
        myVis.YAxis.Expression = 'Avg(if([Measure]="Trend_RXQUAL UL Average (GSM)",[MeasureValue])) as [RXQUAL UL Average (GSM)], Avg(if([Measure]="Trend_RXQUAL DL Average (GSM)",[MeasureValue])) as [RXQUAL DL Average (GSM)]'
        tableVis.Data.WhereClauseExpression = '([Measure] = "Trend_RXQUAL UL Average (GSM)" or [Measure] = "Trend_RXQUAL DL Average (GSM)") and [CellName] = "'+selected_cell+'"'
    if selected_measure == "FER UL & DL Average (GSM)":
        myVis.YAxis.Expression = 'Avg(if([Measure]="Trend_FER UL Average (GSM)",[MeasureValue])) as [FER DL Average (GSM)], Avg(if([Measure]="Trend_FER DL Average (GSM)",[MeasureValue])) as [FER DL Average (GSM)]'
        tableVis.Data.WhereClauseExpression = '([Measure] = "Trend_FER UL Average (GSM)" or [Measure] = "Trend_FER DL Average (GSM)") and [CellName] = "'+selected_cell+'"'
    if selected_measure == "RXLEV UL & DL Average (dBm)":
        myVis.YAxis.Expression = 'Avg(if([Measure]="Trend_RXLEV UL Average (dBm)",[MeasureValue])) as [RXLEV UL Average (dBm)], Avg(if([Measure]="Trend_RXLEV DL Average (dBm)",[MeasureValue])) as [RXLEV DL Average (dBm)]'
        tableVis.Data.WhereClauseExpression = '([Measure] = "Trend_RXLEV UL Average (dBm)" or [Measure] = "Trend_RXLEV DL Average (dBm)") and [CellName] = "'+selected_cell+'"'
else:
    myVis.Data.WhereClauseExpression = '[Measure] = "Trend_${trendDropDown}" and [CellName] = "'+selected_cell+'"'
    tableVis.Data.WhereClauseExpression = '[Measure] = "Trend_${trendDropDown}" and [CellName] = "'+selected_cell+'"'
    myVis.YAxis.Expression = 'Avg([MeasureValue]) as ['+selected_measure+']'


if selected_cell == "(All)":
    tableVis.Data.WhereClauseExpression = tableVis.Data.WhereClauseExpression.split(" and [CellName]")[0]
    if "UL & DL" in selected_measure:
       myVis.Data.WhereClauseExpression = myVis.Data.WhereClauseExpression.split("[CellName]")[0]
    else:
       myVis.Data.WhereClauseExpression = myVis.Data.WhereClauseExpression.split(" and [CellName]")[0]


