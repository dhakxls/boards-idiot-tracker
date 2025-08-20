Attribute VB_Name = "Module1"
' Module1 (Standard Module)
Sub SortAmountColumns(ws As Worksheet)
    Dim lastColumn As Long
    Dim currentCol As Long
    Dim amountFound As Boolean

    ' Find the last column with data in row 2
    lastColumn = ws.Cells(2, ws.Columns.Count).End(xlToLeft).Column

    ' Loop through each column in row 2 to find and sort the "Amount" columns
    For currentCol = 1 To lastColumn
        ' Check if the current cell in row 2 contains "Amount"
        If ws.Cells(2, currentCol).value = "Amount" Then
            ' Determine the range for sorting based on the current "Amount" column
            Dim startCell As Range
            Dim endCell As Range

            Set startCell = ws.Cells(3, currentCol - 3)
            Set endCell = ws.Cells(ws.Rows.Count, currentCol).End(xlUp).Offset(1, 0)

            ' Sort the data range starting from row 3 and including the previous two columns and the "Amount" column
            SortDataRange ws.Range(startCell, endCell)
            amountFound = True
        End If
    Next currentCol

    ' If no "Amount" column is found, display a message
    If Not amountFound Then
        MsgBox "No 'Amount' column found in row 2.", vbExclamation
    End If
End Sub

Sub SortDataRange(rng As Range)
    ' Sort the range based on the values in the "Amount" column in descending order
    With rng
        .Sort Key1:=.Columns(4), Order1:=xlDescending, Header:=xlNo
    End With
End Sub
