VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub CheckMatch()
    Dim shtColl As New Collection
    Dim i As Long
    
'add the list into a collection
    With Workbooks.Open(ThisWorkbook.Path & "\Reference\FileA.xlsx")
        .Activate
            With .Worksheets("list")
                Set startingCell = .Range("A1")
                For Each cell In .Range(startingCell, startingCell.End(xlDown))
                    shtColl.Add cell.Text, cell
                Next
            End With
        .Close
    End With
    
'delete the item in collection if the sheet exists in DCB
    With Workbooks.Open(ThisWorkbook.Path & "\Reference\FileB.xlsx")
        .Activate
        For Each sht In .Sheets
            shtColl.Remove sht.Name
        Next
        .Close
        
    'then display the revised collection as msgbox or debug.print
        For Each Item In shtColl
            Debug.Print Item
        Next
    End With

End Sub
