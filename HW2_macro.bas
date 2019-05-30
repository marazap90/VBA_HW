Attribute VB_Name = "Module2"
Sub Tickerz():

    Dim LastRow As Double
    Dim LastResultRow As Double
    LastResultRow = 2
    
    Dim Total As Double
    Dim Tick As String
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Total Stock Volume"
    
    Dim YrOpen As Double
    Dim YrClose As Double
    Dim Diff As Double
    Dim Pct As Double
    Cells(1, 12).Value = "Year Open"
    Cells(1, 13).Value = "Year Close"
    Cells(1, 14).Value = "Yearly Change"
    Cells(1, 15).Value = "Yearly Pct Change"
      
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
    
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
        
            Tick = Cells(i, 1).Value
            Total = Cells(i, 7).Value
            YrOpen = Cells(i, 3).Value
        
            Range("I" & LastResultRow).Value = Tick
            Range("L" & LastResultRow).Value = YrOpen
        
        Else
            Total = Total + Cells(i, 7).Value
            
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                Range("J" & LastResultRow).Value = Total
                    
                YrClose = Cells(i, 6).Value
                Range("M" & LastResultRow).Value = YrClose
                Diff = YrClose - YrOpen
                Range("N" & LastResultRow).Value = Diff
                    
                        If YrOpen = 0 Then
                           Range("O" & LastResultRow).Value = "N/A"
                        Else
                            Pct = Diff / YrOpen
                            Range("O" & LastResultRow).Value = Pct
                        End If
            
                LastResultRow = LastResultRow + 1
                
            End If
              
        End If
            
    Next i
    
    Range("O:O").NumberFormat = "0.00%"

End Sub


