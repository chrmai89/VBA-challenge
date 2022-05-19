Sub Ticker()

Dim i As Double
Dim TickerVol As Double
Dim SummaryTable As Integer
Dim Opn As Double
Dim Cls As Double

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

TickerVol = 0
SummaryTable = 2

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow

    If Cells(i, 1).Value = Cells(i + 1, 1).Value And Cells(i, 1).Value = Cells(i - 1, 1).Value Then
        TickerVol = TickerVol + Cells(i, 7)
    
    ElseIf Cells(i, 1).Value = Cells(i + 1, 1).Value And Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
        TickerVol = TickerVol + Cells(i, 7)
        Opn = Cells(i, 3).Value
        
    ElseIf Cells(i, 1).Value = Cells(i - 1, 1).Value And Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        TickerVol = TickerVol + Cells(i, 7)
        Cls = Cells(i, 6).Value
    
        Tickers = Cells(i, 1).Value
        Range("i" & SummaryTable).Value = Tickers
        Range("l" & SummaryTable).Value = TickerVol
        Range("j" & SummaryTable).Value = Cls - Opn
        Range("k" & SummaryTable).Value = FormatPercent((Cls - Opn) / VBA.IIf(Opn = 0, 1, Opn))
        
        SummaryTable = SummaryTable + 1
        TickerVol = 0
        
        Else
        TickerVol = TickerVol + Cells(i, 7).Value
        
        End If
        
    Next i
        MsgBox ("Bingo")
        
        
        
        
YrTable = 10
lastrow = Cells(Rows.Count, YrTable).End(xlUp).Row

For j = 2 To lastrow

    If Cells(j, YrTable).Value >= 0 Then
    
        Cells(j, YrTable).Interior.ColorIndex = 4
        
        Else
        Cells(j, YrTable).Interior.ColorIndex = 3
        
        End If
        
    Next j
    


Dim rng As Range
Dim Maxi As Long
Dim Mini As Long

Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

Set rng = Range("K2:K5000")

Maxi = Application.WorksheetFunction.Max(rng)
Mini = Application.WorksheetFunction.Min(rng)


For x = 2 To 5000

    If Cells(x, 11).Value = Maxi Then
    
    Cells(2, 17).Value = Maxi
    Cells(2, 16).Value = Cells(x, 9).Value
    
    ElseIf Cells(x, 11).Value = Mini Then
    
    Cells(3, 17).Value = Mini
    Cells(3, 16).Value = Cells(x, 9).Value
    
    
    End If
    
    Next x
    
    'I was able to get the Greatest % but couldn't format the Value.
    'It only worked for the 2014 worksheet.
    
    
    
Dim rngVol As Range
Dim MaxTol As Long
Set rngVol = Range("L2:L3000")
        
MaxTol = Application.WorksheetFunction.Max(rngVol)
        
    For y = 2 To 3000
        If Cells(y, 12).Value = MaxTol Then
            Cells(4, 17).Value = MaxTol
            Cells(4, 16).Value = Cells(y, 9)
            
            End If
            
            Next y
        
      'Sorry I was unable to get the Greatest Total Value to work.
      'I tried the above but couldn't figure how to fix the debug
        
        
        
End Sub