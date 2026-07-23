Attribute VB_Name = "Module1"
Sub REPORT_MAIN1()
    Dim row As Integer
    Dim col As Integer
    Dim ws As Worksheet

    ' Call your custom date input sub
    Call Ddate(2) ' 2 = prompts for start & end dates

    ' If user cancels or leaves input empty
    If Sdat = "" Or Edat = "" Then
        MsgBox "Start or End date not provided. Exiting.", vbExclamation, "Missing Date"
        Exit Sub
    End If

    ' Set target worksheet from active workbook
    On Error Resume Next
    Set ws = ActiveWorkbook.Sheets("Sheet1")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Sheet 'Sheet1' not found in the active workbook.", vbCritical, "Missing Sheet"
        Exit Sub
    End If

    ' Optional: Clear previous data from known range
    ws.Range("C17:E27").ClearContents

    ' Connect to SQL (Assumed to be managed by Add-In)
    Connect

    ' Ensure conn and rst are initialized
    If conn Is Nothing Then
        MsgBox "Connection not initialized.", vbCritical, "Connection Error"
        Exit Sub
    End If

    If rst Is Nothing Then
        MsgBox "Recordset object not initialized.", vbCritical, "Recordset Error"
        Exit Sub
    End If

    ' Execute stored procedure with user-defined dates
    On Error GoTo ErrHandler
    rst.Open "EXEC Stored procedure '" & Sdat & "', '" & Edat & "'", conn

    ' Check if data exists
    If rst.EOF Then
        MsgBox "No record for selected date range", vbExclamation, "No Data"
        GoTo Cleanup
    End If

    ' Define cell mappings for the info you want in my case "VALUE" and "VOLUME"
    Dim cellMappings As Variant
    cellMappings = Array("C17", "E17", "C18", "E18", "C22", "E22", "C23", "E23", "C27", "E27")

    Dim i As Integer
    i = 0

    ' Loop through recordsets and insert "VALUE" and "VOLUME"
    Do
        If Not rst.EOF Then
            Do While Not rst.EOF
                If i <= UBound(cellMappings) - 1 Then
                    ' VALUE
                    If Not IsNull(rst.Fields(1).Value) Then
                        ws.Range(cellMappings(i)).Value = rst.Fields(1).Value
                    Else
                        ws.Range(cellMappings(i)).Value = 0
                    End If
                    
                    ' VOLUME
                    If Not IsNull(rst.Fields(2).Value) Then
                        ws.Range(cellMappings(i + 1)).Value = rst.Fields(2).Value
                    Else
                        ws.Range(cellMappings(i + 1)).Value = 0
                    End If
                    
                    i = i + 2
                End If
                rst.MoveNext
            Loop
        End If

        ' Move to next recordset if exists
        Set rst = rst.NextRecordset
        If rst Is Nothing Then Exit Do
    Loop

Cleanup:
    ' Safely disconnect and clean up
    On Error Resume Next
    Disconnect
    ws.Range("A1").Select
    MsgBox "E-Levy data inserted successfully into 'Sheet1'!", vbInformation, "Success"
    Exit Sub

ErrHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Runtime Error"
    Resume Cleanup
End Sub
