Attribute VB_Name = "Module1"
Sub FindEmployeesWithLessThan10HoursBetweenShiftsAndSaveToFile()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim timeBetweenShifts As Date
    Dim employeeName As String
    Dim previousEndTime As Date
    Dim outputText As String
    
    ' Change the sheet name to match your worksheet name
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Initialize variables
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    outputText = ""
    
    ' Loop through the rows
    For i = 2 To lastRow ' Assuming the data starts from row 2 (with headers)
        Dim startTime As Variant
        Dim endTime As Variant
        Dim timeDiff As Double
        
        startTime = ws.Cells(i, 3).Value ' Assuming "Time" is in column C
        endTime = ws.Cells(i, 4).Value ' Assuming "Time Out" is in column D
        employeeName = ws.Cells(i, 8).Value ' Assuming "Employee Name" is in column H
        
        ' Check if the data is valid date/time values
        If IsDate(startTime) And IsDate(endTime) Then
            ' Calculate the time difference in hours
            timeDiff = DateDiff("h", startTime, endTime)
            
            ' Check if the time difference is less than 10 hours but greater than 1 hour
            If timeDiff < 10 And timeDiff > 1 Then
                ' Check if this is not the first entry for the employee
                If previousEndTime <> 0 Then
                    timeBetweenShifts = DateDiff("h", previousEndTime, startTime)
                    If timeBetweenShifts < 10 And timeBetweenShifts > 1 Then
                        outputText = outputText & employeeName & " has less than 10 hours between shifts: " & Format(startTime, "MM/dd/yyyy hh:mm AM/PM") & " - " & Format(endTime, "MM/dd/yyyy hh:mm AM/PM") & vbCrLf
                    End If
                End If
            End If
            
            ' Update the previous end time for the next iteration
            previousEndTime = endTime
        End If
    Next i
    
    ' Create and write the result to a text file (output2.txt)
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\output2.txt"
    Open filePath For Output As #1
    
    If outputText <> "" Then
        Print #1, "Employees with less than 10 hours between shifts (greater than 1 hour):" & vbCrLf & outputText
    Else
        Print #1, "No employees have less than 10 hours between shifts (greater than 1 hour)."
    End If
    Close #1
End Sub
