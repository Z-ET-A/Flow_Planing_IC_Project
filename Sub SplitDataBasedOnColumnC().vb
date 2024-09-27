Sub SplitDataBasedOnColumnC()
    Dim wsMain As Worksheet
    Dim wsHofer As Worksheet
    Dim wsLowell As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim hoferRow As Long
    Dim lowellRow As Long
    ' Define the main sheet and target sheets
    Set wsMain = ThisWorkbook.Sheets("Active data")
    ' Change to your main sheet name
    Set wsHofer = ThisWorkbook.Sheets("Hofer")
    ' Ensure the sheet exists or create it manually
    Set wsLowell = ThisWorkbook.Sheets("Lowell")
    ' Ensure the sheet exists or create it manually
    ' Initialize row counters for target sheets
    hoferRow = 2
    ' Start at row 2 to account for headers
    lowellRow = 2
    ' Find the last row in column C
    lastRow = wsMain.Cells(wsMain.Rows.Count, "C").End(xlUp).Row
    ' Loop through each row in column C
    For i = 2 To lastRow
    ' Assuming row 1 has headers
        If Not IsError(wsMain.Cells(i, 3)) Then
            If wsMain.Cells(i, 3).Value = "Hofer" Then
            ' Copy the entire row to the Hofer sheet
                wsMain.Rows(i).Copy Destination:=wsHofer.Rows(hoferRow)
                hoferRow = hoferRow + 1
            ElseIf wsMain.Cells(i, 3).Value = "Lowell" Then
            ' Copy the entire row to the Lowell sheet
            wsMain.Rows(i).Copy Destination:=wsLowell.Rows(lowellRow)
            lowellRow = lowellRow + 1
            End If
        End If
    Next i
End Sub