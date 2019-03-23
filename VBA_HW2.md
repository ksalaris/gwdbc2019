Sub volume()

Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet

For Each ws In ThisWorkbook.Worksheets
    ws.Activate

    Dim lrow As Long
    Dim lcol As Long
    lrow = Cells(Rows.Count, 1).End(xlUp).Row
    lcol = Cells(1, Columns.Count).End(xlToLeft).Column
    
    Dim ticker As String
    Dim vol As Double
    
    vol = 0
    summary_row = 2
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    
    

    For i = 2 To lrow
        If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
            vol = vol + Cells(i, 7).Value
            End If
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Cells(summary_row, 9).Value = Cells(i, 1).Value
            Cells(summary_row, 12).Value = vol
            summary_row = summary_row + 1
            vol = 0
            End If
        Next
    
    Dim openprice As Double
    Dim closeprice As Double
    Dim netchange As Double
    Dim percentchange As Double
    summary_row = 2
    
    
    For i = 2 To lrow
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            openprice = Cells(i, 3).Value
            End If
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            If openprice <> 0 Then
                closeprice = Cells(i, 6).Value
                netchange = closeprice - openprice
                percentchange = ((closeprice - openprice) / openprice) * 100
                Cells(summary_row, 10).Value = netchange
                    If netchange > 0 Then Cells(summary_row, 10).Interior.ColorIndex = 4
                    If netchange < 0 Then Cells(summary_row, 10).Interior.ColorIndex = 3
                Cells(summary_row, 11).Value = percentchange
                    If percentchange > 0 Then Cells(summary_row, 11).Interior.ColorIndex = 4
                    If percentchange < 0 Then Cells(summary_row, 11).Interior.ColorIndex = 3
                summary_row = summary_row + 1
                End If
            If openprice = 0 Then
                netchange = closeprice
                Cells(summary_row, 10).Value = netchange
                    If netchange > 0 Then Cells(summary_row, 10).Interior.ColorIndex = 4
                    If netchange < 0 Then Cells(summary_row, 10).Interior.ColorIndex = 3
                Cells(summary_row, 11).Value = "N/A"
                summary_row = summary_row + 1
                End If
                
        End If
        
    Next
      
    Dim maxvol As Double
    Dim maxup As Double
    Dim maxdown As Double
    
    
    maxvol = 0
    maxup = 0
    maxdown = 0
    
      
    For i = 2 To summary_row:
    
        If Cells(i, 11).Value > maxup And Cells(i, 11).Value <> "N/A" Then
            maxup = Cells(i, 11).Value
            Cells(2, 16).Value = Cells(i, 9).Value
            Cells(2, 17).Value = maxup
            End If
        If Cells(i, 11).Value < maxdown Then
            maxdown = Cells(i, 11).Value
            Cells(3, 16).Value = Cells(i, 9).Value
            Cells(3, 17).Value = maxdown
            End If
       If Cells(i, 12).Value > maxvol Then
            maxvol = Cells(i, 12).Value
            Cells(4, 16).Value = Cells(i, 9).Value
            Cells(4, 17).Value = maxvol
            End If
    Next
    
    
Next

    
End Sub

