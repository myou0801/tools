VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Public Function GetTargetWorkerList() As Collection
    Dim lastRow As Long
    Dim i As Long
    Dim TargetRange As Range
    Dim Col As New Collection
    Dim WkBook As WorkerBook

    lastRow = Sheet1.Cells(Sheet1.Rows.Count, "D").End(xlUp).Row

    For i = 2 To lastRow - 1
        Set TargetRange = Sheet1.Cells(i, 4)

        If TargetRange.Value = "ÅZ" Then
            Set WkBook = New WorkerBook
            WkBook.BookPath = TargetRange.Offset(0, -1).Value
            WkBook.WorkerName = TargetRange.Offset(0, -2).Value
            Call Col.Add(WkBook)
        End If
    Next i

    Set GetTargetWorkerList = Col
End Function
