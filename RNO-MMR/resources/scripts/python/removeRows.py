from Spotfire.Dxp.Data import *
table= Document.Data.Tables["Loaded Single Recordings"]
table.RemoveRows(RowSelection(IndexSet(table.RowCount,True))) 