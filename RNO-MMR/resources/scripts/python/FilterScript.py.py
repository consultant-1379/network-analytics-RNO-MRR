from  Spotfire.Dxp.Data  import DataValueCursor
from Spotfire.Dxp.Application.Visuals import *
from Spotfire.Dxp.Application.Filters import *
from Spotfire.Dxp.Data.DataFunctions import DataFunctionExecutorService, DataFunctionInvocation, DataFunctionInvocationBuilder
from Spotfire.Dxp.Data import *
from Spotfire.Dxp.Data import DataFlowBuilder
from Spotfire.Dxp.Application.Filters import ListBoxFilter
from System.Collections.Generic import List
from System.IO import StringReader, StreamReader, StreamWriter, MemoryStream, SeekOrigin
from Spotfire.Dxp.Data.Import import *
myLatCursor = DataValueCursor.CreateFormatted(Document.Data.Tables[ "HistCellChGrColumns" ].Columns[ "CellName" ])


'''markedRows = Document.Data.Markings[ "Marking" ].GetSelection(Document.Data.Tables[ "HistCellChGrColumns" ]).AsIndexSet()
for row in Document.Data.Tables[ "HistCellChGrColumns" ].GetRows(markedRows, myLatCursor):
    print myLatCursor.CurrentValue
    Documents.Properties[ 'overviewHistCellList' ] = myLatCursor.CurrentValue'''
dataTableMap = Document.Data.Tables["HistCellChGrColumns"]
value1="(All)"
for Scheme in Document.FilteringSchemes:
	if Scheme.FilteringSelectionReference.Name == "Filtering scheme":
		print Scheme
		filterElement=Scheme.Item[dataTableMap].Item[dataTableMap.Columns.Item["CellName"]].As[ListBoxFilter]()
		s = filterElement
		for value in filterElement.SelectedValues:
			value1=value
Document.Properties[ 'overviewHistCellList' ] = value1