VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Multiple_year_stock_data()

Cells(1, 10).Value = "Ticker"
Cells(1, 11).Value = "Yearly Change"
Cells(1, 12).Value = "Percent Change"
Cells(1, 13).Value = "Total Stock Volume"

'Setting range and cell data types'
    Dim rng As Range
    Dim cell As Range

'Setting range from A2 to last cell in A column'
    Dim LR As Long
    LR = Cells(Rows.Count, 1).End(xlUp).Row
    
    Set rng = Range("A2:A" & LR)
    
'Defining currentTicker is String, currentTicker not yet operationalized, i will be the changing index as the code runs through'
    Dim currentTicker As String
    currentTicker = ""
    Dim i As Integer
    i = 1

'Defining startValue and endValue as Double; needed for yearly change'
    Dim startValue As Double
    Dim endValue As Double

'Defining totalVol as Long (keep running into issues with overflow); totalVol starts off at 0'
    Dim totalVol As Long
    totalVol = 0

    For Each cell In rng

        Dim cellTicker As String
        Dim cellDate As String
        Dim cellOpen As Double
        Dim cellHigh As Double
        Dim cellLow As Double
        Dim cellClose As Double
        Dim cellVol As Long

'Defining WHERE each value is'
        cellTicker = Cells(cell.Row, cell.Column).Value
        cellDate = Cells(cell.Row, cell.Column + 1).Value
        cellOpen = Cells(cell.Row, cell.Column + 2).Value
        cellHigh = Cells(cell.Row, cell.Column + 3).Value
        cellLow = Cells(cell.Row, cell.Column + 4).Value
        cellClose = Cells(cell.Row, cell.Column + 5).Value
        cellVol = Cells(cell.Row, cell.Column + 6).Value

        If StrComp(currentTicker, cellTicker) Then
            currentTicker = cellTicker
            i = i + 1

            'ticker
            Cells(i, 10).Value = currentTicker

            If startValue Then
            endValue = Cells(cell.Row - 1, cell.Column + 5).Value

            'yearly change
            Dim yearlyChange As Double
            yearlyChange = endValue - startValue
            Cells(i - 1, 11).Value = yearlyChange

            If yearlyChange > 0 Then
                Cells(i - 1, 11).Interior.ColorIndex = 4
            Else
                Cells(i - 1, 11).Interior.ColorIndex = 3
            End If

            'percent change
            Cells(i - 1, 12).Value = (yearlyChange / startValue)
            Cells(i - 1, 12).NumberFormat = "0.00%"

            'total stock vol
            Cells(i - 1, 13).Value = totalVol
            totalVol = totalVol & currentVol

            End If

            startValue = cellOpen

        End If

    Next cell

End Sub

