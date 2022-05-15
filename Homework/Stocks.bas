Attribute VB_Name = "Module1"
Sub Stocks()

For Each ws In Worksheets
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim Ticker As String
Dim Row As Integer
Dim Yearly_Change As Double
Dim OpenSum As Double
Dim CloseSum As Double
Dim Percent As Single
Dim StockVolume As Double



Percent = 0
Row = 2
OpenSum = 0
CloseSum = 0
StockVolume = 0

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("N2").Value = "Greatest Increase"
ws.Range("N3").Value = "Greatest Decrease"
ws.Range("N4").Value = "Greatest Total Volume"

For i = 2 To lastrow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker = ws.Cells(i, 1).Value
        OpenSum = OpenSum + ws.Cells(i, 3).Value
        CloseSum = CloseSum + ws.Cells(i, 6).Value
        StockVolume = StockVolume + ws.Cells(i, 7).Value
        ws.Range("I" & Row).Value = Ticker
        ws.Range("J" & Row).Value = CloseSum - OpenSum
            If ws.Range("J" & Row).Value >= 0 Then
                ws.Range("J" & Row).Interior.ColorIndex = 4
            Else
                ws.Range("J" & Row).Interior.ColorIndex = 3
            End If
            
            If OpenSum <> 0 Then
                Percent = (ws.Range("J" & Row).Value / OpenSum) * 100
            End If
    
        ws.Range("K" & Row).Value = Percent & "%"
        ws.Range("L" & Row).Value = StockVolume
        
        
        Row = Row + 1
        OpenSum = 0
        CloseSum = 0
        Percent = 0
        StockVolume = 0
        
      
        
    Else
        OpenSum = OpenSum + ws.Cells(i, 3).Value
        CloseSum = CloseSum + ws.Cells(i, 6).Value
        StockVolume = StockVolume + ws.Cells(i, 7).Value
    End If
    

 Next i
 
 ws.Range("O2").Value = WorksheetFunction.Max(ws.Range("J2:J" & lastrow))
 ws.Range("O3").Value = WorksheetFunction.Min(ws.Range("J2:J" & lastrow))
 ws.Range("O4").Value = WorksheetFunction.Max(ws.Range("L2:L" & lastrow))
 
Next ws

End Sub

