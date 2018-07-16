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
    Dim sht As Worksheet
    Dim shtColl As New Collection
    Dim FileA As Workbook
    Dim FileB As Workbook
    Dim i As Long
    
'add the list into a collection
    Set FileA = Workbooks.Open(ThisWorkbook.Path & "\Reference\FileA.xlsx")
    
    FileA.Activate
With FileA.Sheets("list")
    i = 1
    Do Until .Range("A" & i) = ""
        shtColl.Add .Range("A" & i)
        i = i + 1
    Loop
End With

'delete the item in collection if the sheet exists in DCB
    Set FileB = Workbooks.Open(ThisWorkbook.Path & "\Reference\FileB.xlsx")
    
    FileB.Activate
    
    For i = 1 To shtColl.Count
        For Each sht In FileB.Sheets
            If sht.Name = shtColl(i) Then

            shtColl.Remove (i)
            End If
        Next sht
    Next i
    
'then display the revised collection as msgbox or debug.print
    For i = 1 To shtColl.Count
    Debug.Print shtColl(i)
    Next i

End Sub
