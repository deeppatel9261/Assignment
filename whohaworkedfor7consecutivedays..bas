Attribute VB_Name = "Module3"
Sub FindEmployeesWith7ConsecutiveDaysAndSaveToFile()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim consecutiveDays As Long
    Dim consecutiveFlag As Boolean
    
    ' Change the sheet name to match your worksheet name
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Initialize variables
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    consecutiveDays = 0
    consecutiveFlag = False
    
    ' Store the unique days worked by employees in a dictionary
    Dim workDays As Object
    Set workDays = CreateObject("Scripting.Dictionary")
    
    ' Store information about employees who worked for 7 consecutive days
    Dim employees7ConsecutiveDays As String
    employees7ConsecutiveDays = ""
    
    ' Loop through the rows
    For i = 2 To lastRow ' Assuming the data starts from row 2 (with headers)
        Dim startDate As Variant
        Dim employeeName As String
        
        startDate = ws.Cells(i, 6).Value ' Assuming "Pay Cycle Start Date" is in column F
        employeeName = ws.Cells(i, 8).Value ' Assuming "Employee Name" is in column H
        
        ' Check if the data is a valid date
        If IsDate(startDate) Then
            If Not workDays.Exists(employeeName & "-" & Format(startDate, "MM/dd/yyyy")) Then
                workDays(employeeName & "-" & Format(startDate, "MM/dd/yyyy")) = True
                consecutiveDays = consecutiveDays + 1
                employees7ConsecutiveDays = employees7ConsecutiveDays & employeeName & " worked on " & Format(startDate, "MM/dd/yyyy") & vbCrLf
            End If
        End If
        
        If consecutiveDays = 7 Then
            consecutiveFlag = True
            Exit For ' No need to continue checking
        End If
    Next i
    
    ' Create and write the result to a text file (output.txt)
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\output.txt"
    Open filePath For Output As #1
    
    If consecutiveFlag Then
        Print #1, "Employees who have worked for 7 consecutive days:" & vbCrLf & employees7ConsecutiveDays
    Else
        Print #1, "No employees worked for 7 consecutive days."
    End If
    Close #1
End Sub

