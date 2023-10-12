Attribute VB_Name = "Module2"
Sub FindEmployeesWithMoreThan14HoursInSingleShiftAndSaveToFile()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim employeeName As String
    Dim shiftDuration As Double
    Dim outputText As String
    
    ' Change the sheet name to match your worksheet name
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Initialize variables
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    outputText = ""
    
    ' Loop through the rows
    For i = 2 To lastRow ' Assuming the data starts from row 2 (with headers)
        Dim startTime As Date
        Dim endTime As Date
        
        startTime = ws.Cells(i, 3).Value ' Assuming "Time" is in column C
        endTime = ws.Cells(i, 4).Value ' Assuming "Time Out" is in column D
        employeeName = ws.Cells(i, 8).Value ' Assuming "Employee Name" is in column H
        
        ' Check if the data is valid date/time values
        If IsDate(startTime) And IsDate(endTime) Then
            ' Calculate the shift duration in hours
            shiftDuration = DateDiff("h", startTime, endTime)
            
            ' Check if the shift duration is more than 14 hours
            If shiftDuration > 14 Then
                outputText = outputText & employeeName & " worked for more than 14 hours in a single shift: " & Format(startTime, "MM/dd/yyyy hh:mm AM/PM") & " - " & Format(endTime, "MM/dd/yyyy hh:mm AM/PM") & vbCrLf
            End If
        End If
    Next i
    
    ' Create and write the result to a text file (output3.txt)
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\output3.txt"
    Open filePath For Output As #1
    
    If outputText <> "" Then
        Print #1, "Employees who have worked for more than 14 hours in a single shift:" & vbCrLf & outputText
    Else
        Print #1, "No employees have worked for more than 14 hours in a single shift."
    End If
    Close #1
End Sub

