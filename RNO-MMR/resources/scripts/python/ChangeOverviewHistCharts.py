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
# Name    : ChangeOverviewHistCharts.py
# Date    : 01/01/2020
# Revision: 1.0
# Purpose : ChangeOverviewHistCharts
#
# Usage   : NetAn RNO MRR
#


from Spotfire.Dxp.Application.Visuals import ScatterPlot, MarkerType,MarkerShape, CategoryMode, AxisEvaluationMode
from Spotfire.Dxp.Application.Layout import LayoutDefinition
from Spotfire.Dxp.Application.Visuals import *
from Spotfire.Dxp.Data import *


chartDict = {
    'FER UL & DL': ['FER Range', 'FER'],
    'FER UL & DL Acc.': ['FER Range', 'FER'],
    'Path Loss Diff.': ['PLDIFF Range', 'PLDIFF'],
    'Path Loss Diff. Acc.': ['PLDIFF Range', 'PLDIFF'],
    'Path Loss UL & DL': [['PATHLOSS_Min','PATHLOSS_Max'], 'PLOSS'],
    'Path Loss UL & DL Acc.': [['PATHLOSS_Min','PATHLOSS_Max'], 'PLOSS'],
    'Power Lev. MS': ['MSPOWER Range', 'MSPOWER'],
    'Power Lev. MS Acc.': ['MSPOWER Range', 'MSPOWER'],
    'Power Red. BS': ['BSPOWER Range', 'BSPOWER'],
    'Power Red. BS Acc.': ['BSPOWER Range', 'BSPOWER'],
    'RXLEV UL & DL': ['RXLEV Range', 'RXLEV'],
    'RXLEV UL & DL Acc.': ['RXLEV Range', 'RXLEV'],
    'RXQUAL UL & DL': ['RXQUAL Range', 'RXQUAL'],
    'RXQUAL UL & DL Acc.': ['RXQUAL Range', 'RXQUAL'],
    'Timing Adv.': [['TAVAL_Min','TAVAL_Max'], 'TAVAL'],
    'Timing Adv. Acc.': [['TAVAL_Min','TAVAL_Max'], 'TAVAL']
}

yaxisDict = {
    'BSC': '[BSC - % Of Samples]',
    'BSC Acc': '[BSC Accumulated - % of Samples]',
    'Cell': '[Cell - % of Samples]',
    'Cell Acc': '[Cell Accumulated - % of Samples]',
    'Ch_Gr': '[Channel Group - % of Samples]',
    'Ch_Gr Acc': '[Channel Group Accumulated - % of Samples]'
}


def setYaxis(selectLevel, currChart):
    ''' Changes y axis for either accumulated or aboslotue for cell or bsc '''

    if 'Acc.' in currChart:
        yaxisVal = yaxisDict[selectLevel + " Acc"]
    else:
        yaxisVal = yaxisDict[selectLevel]

    return yaxisVal


def setWhereClause(selectLevel, whereClause):
    ''' filters data on chart to selected chart charateristic '''

    if selectLevel == 'BSC':
        whereClause = "[Measurement_Type]~='" + whereClause + "'"
    elif selectLevel == 'Cell':
        whereClause = "[Measurement_Type]~='" + whereClause + "' AND [CellName] = '" + cellSelection + "'"
    else:
        # select lvel is channel group
        # check if cell_selection is all
        if cellSelection == '(All)':
            whereClause = "[Measurement_Type]~='" + whereClause + "' AND [ChannelGroup] = '" + channel_group_selection + "'"
        else:
            whereClause = "[Measurement_Type]~='" + whereClause + "' AND [CellName] = '" + cellSelection + "' AND [ChannelGroup] = '" + channel_group_selection + "'"

    return whereClause


def updateChartDetails(myvis, currMeasureIndex, selectLevel):
    ''' updates selected chart by adding axis, color etc. '''

    # set data table ref
    myvis.Data.DataTableReference = dataTable

    # get current chart from list of kpis selected
    currChart = measureListLimited[currMeasureIndex]

    # set title
    myvis.Title = currChart

    # color by measurement type column
    if singleOverview:
        myvis.ColorAxis.Expression = "Measurement_Type"
    else:
        # set color by type for comparison field
        if(Document.Properties['histogramType'] == 'single_histogram'):
            colorByType = '[FILE_ID]'
        elif (Document.Properties['histogramType'] == 'group_histogram'):
            colorByType = '[GroupingName]'
        myvis.ColorAxis.Expression = """<[Measurement_Type] NEST {0}>""".format(colorByType)

    # get the column to use for the range for x axis
    if currChart in ["Timing Adv.","Path Loss UL & DL", 'Path Loss UL & DL Acc.','Timing Adv. Acc.']:
        chartRange = ' NEST '.join(['['+i+']' for i in chartDict[currChart][0]])
        chartRange = '<' + chartRange + '>'
    else:
        chartRange = '[' + chartDict[currChart][0] + ']'

    myvis.XAxis.Expression = chartRange

    # assign Y axis based on BSC/Node - Accumualted,non accumulated
    myvis.YAxis.Expression = 'Avg(' + setYaxis(selectLevel, currChart) + ')'

    # where clause - default is by measurement type, if cell selected then this is added
    whereClause = chartDict[currChart][1]
    myvis.Data.WhereClauseExpression = setWhereClause(selectLevel, whereClause)

    # assigns the saved color scheme in the analysis file to all points
    if singleOverview:
        myvis.ColorAxis.Coloring.Apply("overviewHistColorScheme")
    else:
        myvis.ColorAxis.Coloring.Clear()
        colorRule = myvis.ColorAxis.Coloring.AddCategoricalColorRule()

    # assign circles as shape
    myvis.ShapeAxis.DefaultShape = MarkerShape(MarkerType.Circle)

    # hide x and y label selectors
    myvis.XAxis.ShowAxisSelector = False
    myvis.YAxis.ShowAxisSelector = False

    # add description for x and y axis, replace acc. from name as not needed for the desc.
    x_axis_desc = currChart.replace("Acc.","")
    myvis.Description = "\nX Axis: " + x_axis_desc + "\nY Axis: % of Samples" 

    # hide legend items excpet for color by
    for i in myvis.Legend.Items:
        if i.Title == "Color by":
            i.Visible = True
            i.ShowTitle = True
            i.ShowAxisSelector = True
        else:
            i.Visible = False  # Hide any legend items not caught above


def leaveOnlyOneBarChart():
    ''' cleans out charts excpet for one anchor chart'''

    for visual in Document.ActivePageReference.Visuals:
        # only delete charts(except for first anchor). Titles on charts with _keep are kept
        if not visual.Title.EndsWith("_KEEP") and not (visual == myChart):
            Document.ActivePageReference.Visuals.Remove(visual)


def addChart(lenMeasureList, currMeasureIndex):
    ''' adds charts to layout'''
    # if 2nd chart or 3rd chart(and if there are 4 charts total) then create new sections for layout
    if (currMeasureIndex == 1 and lenMeasureList != 4) or (currMeasureIndex == 2 and lenMeasureList == 4):
        layout.EndSection()
        layout.BeginStackedSection(10)

    # first chart is already there, so just refernece. Else craete new
    if currMeasureIndex == 0:
        myvis = myChart.As[ScatterPlot]()
    else:
        myvis = Document.ActivePageReference.Visuals.AddNew[ScatterPlot]()

    # call function to update chart format
    updateChartDetails(myvis, currMeasureIndex, selectLevel)

    # add chart to layout
    layout.Add(myvis.Visual, 10)


# ------------
# Main Section
# -------------

# inputs from Script parameters:
#  myTitle - Text Area (Visualization) that will be kept at top
#  visualSettings - Text Area (Visualization) that will hold column seletion
#  dataTable - data table referencing source data for bar charts
#  titlebar = title
#  recordingDetails = recording details at top of charts
#  buttonsArea = action area
#  bottombar = bottom bar with nav buttons

# get current page (this script used for both overview hist and cell comparison chart)
page = Document.ActivePageReference
page_name = page.Title

if page_name == 'MRR Comparison Overview Histogram':
    # get items in listbox and split (stored as a .net list, converting to py list for ease of use)
    measureList = list(Document.Properties["ComparisonOverViewHistList"])
    cellSelection = Document.Properties["ComparisonOverviewHistCellList"]
    channel_group_selection = Document.Properties['comparisonHistogramChannelGroupList']
    singleOverview = False
    error_doc_prop = 'comparisonHistNotification'
else:
    measureList = list(Document.Properties["OverViewHistList"])
    cellSelection = Document.Properties["overviewHistCellList"]
    channel_group_selection = Document.Properties['overviewHistogramChannelGroupList']
    singleOverview = True
    error_doc_prop = 'overviewHistNotification'

# clean out any other charts than the first one
leaveOnlyOneBarChart()

lenMeasureList = len(measureList)

# limit the amount of items to x as we only have x charts
chartLimit = 4
if lenMeasureList >= chartLimit:
    limitData = chartLimit
else:
    limitData = lenMeasureList

# only select first four of selected list
measureListLimited = measureList[:limitData]

# sort the list as the listbox in spotfire is not in sorted order
measureListLimited = sorted(measureListLimited)

# count of limited list
limitedLenMeasureList = len(measureListLimited)

if channel_group_selection == "(All)" and cellSelection == "(All)":
    selectLevel = 'BSC'
elif channel_group_selection != "(All)":
    selectLevel = 'Ch_Gr'
else:
    selectLevel = 'Cell'


# page layout
layout = LayoutDefinition()

# add title
layout.BeginStackedSection()
layout.Add(titleBar, 11)
layout.Add(recordingDetails, 19)

# main container for charts and selction
layout.BeginSideBySideSection(90)

# change first chart to diff proportion as it affects sizing of layout
graphSize = 0

if lenMeasureList == 1:
    graphSize = 20
else:
    graphSize = 10

# laout for first graph
layout.BeginStackedSection(graphSize)

currMeasureIndex = 0
chartCount = 1

# add charts
# loop through the selected list(limited up to 4) and create charts
for currMeasureIndex in range(0, limitedLenMeasureList):
    addChart(limitedLenMeasureList, currMeasureIndex)

# if more selected than 4, give error mesage
if lenMeasureList > chartLimit:
    Document.Properties[error_doc_prop] = 'More than 4 measures selected. Only showing first 4.'
else:
    Document.Properties[error_doc_prop] = ''

# end of add charts
layout.EndSection()

# settings bar
layout.BeginStackedSection(4)
layout.Add(visualSettings, 20)
layout.Add(buttonsArea, 5)
layout.EndSection()

# end main layout
layout.EndSection()

# add bottom bar below everything else
layout.Add(bottomBar, 10)

# end all layouts
layout.EndSection()

# apply layout to screen
page.ApplyLayout(layout)
