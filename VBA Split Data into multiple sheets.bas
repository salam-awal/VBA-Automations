Attribute VB_Name = "Module1"
Sub SplitDataIntoSheets()
    Dim ws As Worksheet
    Dim wsNew As Worksheet
    Dim rng As Range
    Dim branchCol As String
    Dim lastRow As Long
    Dim lastCol As Long
    Dim branch As Variant
    Dim dict As Object
    Dim i As Long
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change to your data sheet name
    branchCol = "C" ' Change to the column letter where branch names are
    lastRow = ws.Cells(ws.Rows.Count, branchCol).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Set rng = ws.Range(branchCol & "2:" & branchCol & lastRow) ' Assuming headers in row 1
    Set dict = CreateObject("Scripting.Dictionary")
    ' Get unique branch names
    For i = 2 To lastRow
        branch = ws.Cells(i, branchCol).Value
        If Not dict.Exists(branch) Then
            dict.Add branch, 1
        End If
    Next i
    ' Copy headers
    Dim headerRange As Range
    Set headerRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
    ' Create sheets and copy data
    Dim br As Variant
    For Each br In dict.Keys
        On Error Resume Next
        Application.DisplayAlerts = False
        Worksheets(br).Delete ' Delete if sheet already exists
        Application.DisplayAlerts = True
        On Error GoTo 0
        Set wsNew = Worksheets.Add(After:=Worksheets(Worksheets.Count))
        wsNew.Name = br
        headerRange.Copy Destination:=wsNew.Range("A1")
        ' AutoFilter and copy filtered data
        ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).AutoFilter Field:=Range(branchCol & "1").Column, Criteria1:=br
        ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).SpecialCells(xlCellTypeVisible).Copy Destination:=wsNew.Range("A2")
        ws.AutoFilterMode = False
    Next br
End Sub
 

