Sub CompareSheets()
    Dim wsLastMonth As Worksheet
    Dim wsThisMonth As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRowLastMonth As Long
    Dim lastRowThisMonth As Long
    Dim lastRowOutput As Long
    Dim rowLastMonth As Range
    Dim rowThisMonth As Range
    Dim idLastMonth As Range
    Dim idThisMonth As Range
    Dim matchFound As Range
    Dim colCount As Integer
    Dim i As Integer
    Dim j As Integer
    
    ' Set worksheets
    Set wsLastMonth = ThisWorkbook.Sheets("Last Month")
    Set wsThisMonth = ThisWorkbook.Sheets("This Month")
    
    ' Create Output sheet if it does not exist, or clear it if it does
    On Error Resume Next
    Set wsOutput = ThisWorkbook.Sheets("Output")
    On Error GoTo 0
    If wsOutput Is Nothing Then
        Set wsOutput = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsOutput.Name = "Output"
    Else
        wsOutput.Cells.Clear
    End If
    
    ' Determine last rows
    lastRowLastMonth = wsLastMonth.Cells(wsLastMonth.Rows.Count, 1).End(xlUp).Row
    lastRowThisMonth = wsThisMonth.Cells(wsThisMonth.Rows.Count, 1).End(xlUp).Row
    
    ' Copy headers to Output sheet
    wsOutput.Cells(1, 1).Value = "Type"
    wsLastMonth.Range("A1:DD1").Copy Destination:=wsOutput.Cells(1, 2)
    
    ' Compare sheets
    colCount = wsLastMonth.Columns.Count
    
    ' Loop through Last Month sheet
    For Each rowLastMonth In wsLastMonth.Range("A2:A" & lastRowLastMonth).Rows
        Set idLastMonth = rowLastMonth.Cells(1, 1)
        Set matchFound = wsThisMonth.Range("A:A").Find(What:=idLastMonth.Value, LookIn:=xlValues, LookAt:=xlWhole)
        
        If matchFound Is Nothing Then
            ' Row deleted
            lastRowOutput = wsOutput.Cells(wsOutput.Rows.Count, 1).End(xlUp).Row + 1
            rowLastMonth.EntireRow.Copy Destination:=wsOutput.Cells(lastRowOutput, 1)
            wsOutput.Cells(lastRowOutput, 1).Insert Shift:=xlToRight
            wsOutput.Cells(lastRowOutput, 1).Value = "Delete"
        Else
            ' Check for changes
            Set rowThisMonth = matchFound.EntireRow
            For i = 2 To colCount
                If rowLastMonth.Cells(1, i).Value <> rowThisMonth.Cells(1, i).Value Then
                    lastRowOutput = wsOutput.Cells(wsOutput.Rows.Count, 1).End(xlUp).Row + 1
                    rowThisMonth.EntireRow.Copy Destination:=wsOutput.Cells(lastRowOutput, 1)
                    wsOutput.Cells(lastRowOutput, 1).Insert Shift:=xlToRight
                    wsOutput.Cells(lastRowOutput, 1).Value = "Change"
                    ' Highlight changes
                    For j = 2 To colCount
                        If rowLastMonth.Cells(1, j).Value <> rowThisMonth.Cells(1, j).Value Then
                            wsOutput.Cells(lastRowOutput, j + 1).Interior.Color = RGB(255, 255, 0)
                        End If
                    Next j
                    Exit For
                End If
            Next i
        End If
    Next rowLastMonth
    
    ' Loop through This Month sheet for added rows
    For Each rowThisMonth In wsThisMonth.Range("A2:A" & lastRowThisMonth).Rows
        Set idThisMonth = rowThisMonth.Cells(1, 1)
        Set matchFound = wsLastMonth.Range("A:A").Find(What:=idThisMonth.Value, LookIn:=xlValues, LookAt:=xlWhole)
        
        If matchFound Is Nothing Then
            ' Row added
            lastRowOutput = wsOutput.Cells(wsOutput.Rows.Count, 1).End(xlUp).Row + 1
            rowThisMonth.EntireRow.Copy Destination:=wsOutput.Cells(lastRowOutput, 1)
            wsOutput.Cells(lastRowOutput, 1).Insert Shift:=xlToRight
            wsOutput.Cells(lastRowOutput, 1).Value = "Added"
        End If
    Next rowThisMonth
    
    ' Autofit columns in Output sheet
    wsOutput.Columns.AutoFit
    
    MsgBox "Comparison complete!", vbInformation
End Sub

