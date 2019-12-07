Sub Stock()
  ' LOOP THROUGH ALL SHEETS
  Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
        ' CLEAR COLUMNS I TO Q
        Columns("I:Q").Select
        Selection.Clear
        ' Add Heading for summary
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Value"
        'Create Variable to hold Value
        Dim Ticker As String
        Dim Stock_Total As Double
        Dim Year_open As Double
        Dim Year_close As Double
        Dim Yearly_change As Double
        Dim Percent_change As Double
        Dim Row As Double
        Dim Column As Integer
        Dim OpenPrice As Variant
        Dim ClosingPrice As Variant
        Dim i As Long

        Stock_Total = 0
        Row = 2
        Column = 1
        'Set Initial Open Price
        OpenPrice = Cells(2, Column + 2).Value
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row

        ' Loop through all ticker symbol
        For i = 2 To LastRow
          ' Check if we are on the same ticker, if not, conditional.......
          If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
            ' Set Ticker name
            Ticker = Cells(i, Column).Value
            Cells(Row, Column + 8).Value = Ticker
            ' Add Total Stock Volume
            Stock_Total = Stock_Total + Cells(i, Column + 6).Value
            Cells(Row, Column + 11).Value = Stock_Total
            ' Set Close Price
            ClosingPrice = Cells(i, Column + 5).Value
            ' Add Yearly Change
            Yearly_change = ClosingPrice - OpenPrice
            Cells(Row, Column + 9).Value = Yearly_Change

            If (OpenPrice = 0 And ClosingPrice = 0) Then
              Percent_Change = 0
            ElseIf (OpenPrice = 0 And ClosingPrice <> 0) Then
              Percent_Change = 1
            Else
              Percent_Change = Yearly_Change / OpenPrice
              Cells(Row, Column + 10).Value = Percent_Change
              Cells(Row, Column + 10).NumberFormat = "0.00%"
            End If
            Row = Row + 1
            OpenPrice = Cells(i + 1, Column + 2)
            Stock_Total = 0
          Else
            Stock_Total = Stock_Total + Cells(i, Column + 6).Value
          End If
        Next i

        ' Estimate Last Row in each WorkSheet
        Yearly_CLR = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row
        ' Change the Cell Colors
        For j = 2 To Yearly_CLR
            If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
                Cells(j, Column + 9).Interior.ColorIndex = 10
            ElseIf Cells(j, Column + 9).Value < 0 Then
                Cells(j, Column + 9).Interior.ColorIndex = 3
            End If
        Next j

        ' Set Greatest Percent Increase, Percent Decrease, and Total Stock Volume
        Cells(2, Column + 14).Value = "Greatest % Increase"
        Cells(3, Column + 14).Value = "Greatest % Decrease"
        Cells(4, Column + 14).Value = "Greatest Total Volume"
        Cells(1, Column + 15).Value = "Ticker"
        Cells(1, Column + 16).Value = "Value"
        ' Look through each rows to find the greatest value and its associate ticker
        For k = 2 To Yearly_CLR
            If Cells(k, Column + 10).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & Yearly_CLR)) Then
                Cells(2, Column + 15).Value = Cells(k, Column + 8).Value
                Cells(2, Column + 16).Value = Cells(k, Column + 10).Value
                Cells(2, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(k, Column + 10).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & Yearly_CLR)) Then
                Cells(3, Column + 15).Value = Cells(k, Column + 8).Value
                Cells(3, Column + 16).Value = Cells(k, Column + 10).Value
                Cells(3, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(k, Column + 11).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & Yearly_CLR)) Then
                Cells(4, Column + 15).Value = Cells(k, Column + 8).Value
                Cells(4, Column + 16).Value = Cells(k, Column + 11).Value
            End If
        Next k
        Columns("I:Q").EntireColumn.AutoFit
      Next ws
End Sub