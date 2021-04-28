VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub VBAChallenge()
Attribute VBAChallenge.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
    
    'Loop End
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Headers
    Range("H1").Value = "Ticker"
    Range("I1").Value = "Yearly Change"
    Range("J1").Value = "Percent Change"
    Range("K1").Value = "Total Stock Volume"
    
    'Variables Declaration
        Dim OpeningPrice As Double
        Dim ClosingPrice As Double
        Dim YearlyChange As Double
        Dim TickerName As String
        Dim PercentChange As Double
        Dim Volume As Double
        Volume = 0
        Dim Row As Double
        Row = 2
        Dim Column As Integer
        Column = 1
        Dim h As Long

    'For <open>
    OpeningPrice = Cells(2, Column + 2).Value
    
    'Ticker Symbol Loop
    For h = 2 To LastRow
    
    'Ticker Symbol Check
    If Cells(h + 1, Column).Value <> Cells(h, Column).Value Then
                
        'Ticker
                TickerName = Cells(h, Column).Value
                Cells(Row, Column + 7).Value = TickerName
                
        ' Close Price
                ClosingPrice = Cells(h, Column + 4).Value
                
        ' Yearly Change
                YearlyChange = ClosingPrice - OpeningPrice
                Cells(Row, Column + 8).Value = YearlyChange
                
        ' Percent Change
                If (OpeningPrice = 0 And ClosingPrice = 0) Then
                    PercentChange = 0
                ElseIf (OpeningPrice = 0 And ClosingPrice <> 0) Then
                    PercentChange = 1
                Else
                    Percent_Change = YearlyChange / OpeningPrice
                    Cells(Row, Column + 9).Value = PercentChange
                    Cells(Row, Column + 9).NumberFormat = "0.00%"
                End If
                
        'Total Volume
                Volume = Volume + Cells(h, Column + 5).Value
                Cells(Row, Column + 10).Value = Volume
                
        'Summary table row
                Row = Row + 1
                
        'Reset Open Price
                OpeningPrice = Cells(h + 1, Column + 1)
                
        'reset the Volumn Total
                Volume = 0
            
        'Same ticker
            Else
                Volume = Volume + Cells(h, Column + 5).Value
            End If
        Next h
        
             'Last Row of Yearly Change per WS
        YCLastRow = ws.Cells(Rows.Count, Column + 7).End(xlUp).Row
        
        ' Set the Cell Colors
            For i = 2 To YCLastRow
            If (Cells(i, Column + 8).Value > 0 Or Cells(i, Column + 8).Value = 0) Then
                Cells(i, Column + 8).Interior.Color = vbGreen
            ElseIf Cells(i, Column + 8).Value < 0 Then
                Cells(i, Column + 8).Interior.Color = vbRed
            End If
        Next i
        
        ' Set Greatest % Increase, % Decrease, and Total Volume
        Cells(2, Column + 13).Value = "Greatest % Increase"
        Cells(3, Column + 13).Value = "Greatest % Decrease"
        Cells(4, Column + 13).Value = "Greatest Total Volume"
        Cells(1, Column + 14).Value = "Ticker"
        Cells(1, Column + 15).Value = "Value"
        ' Look through each rows to find the greatest value and its     associate ticker
        For Z = 2 To YCLastRow
            If Cells(Z, Column + 9).Value = Application.WorksheetFunction.Max(ws.Range("J2:J" & YCLastRow)) Then
                Cells(2, Column + 14).Value = Cells(Z, Column + 7).Value
                Cells(2, Column + 15).Value = Cells(Z, Column + 9).Value
                Cells(2, Column + 15).NumberFormat = "0.00%"
            ElseIf Cells(Z, Column + 10).Value = Application.WorksheetFunction.Min(ws.Range("J2:J" & YCLastRow)) Then
                Cells(3, Column + 14).Value = Cells(Z, Column + 7).Value
                Cells(3, Column + 15).Value = Cells(Z, Column + 9).Value
                Cells(3, Column + 15).NumberFormat = "0.00%"
            ElseIf Cells(Z, Column + 10).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & YCLastRow)) Then
                Cells(4, Column + 14).Value = Cells(Z, Column + 7).Value
                Cells(4, Column + 15).Value = Cells(Z, Column + 10).Value
            End If
        Next Z
        
    Next ws
    


End Sub
