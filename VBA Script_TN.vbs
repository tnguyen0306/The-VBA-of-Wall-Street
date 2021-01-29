Sub VBAchallenge()
    'assign all varable
    Dim lastrow As Double
    Dim lastcol As Double
    Dim ticker As String
    Dim yearlychange As Double
    Dim percentchange As Double
    Dim totalvol As Double
    Dim opencost As Double
    Dim closecost As Double
    Dim tablerow As Double
    Dim greatestPer As Double
    Dim smallestPer As Double
    Dim greatestTotal As Double
    
    'For loop to work on each sheet
    For Each ws In Worksheets
        ws.Activate
        
        'identify Last Row & Column & summary table row
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        lastcol = Cells(1, Columns.Count).End(xlToLeft).Column
        tablerow = 2
        
        'Set up the sheet
        Cells(1, lastcol + 2).Value = "Ticker"
        Cells(1, lastcol + 3).Value = "Yearly Change"
        Cells(1, lastcol + 4).Value = "Percent Change"
        Range("K2:K1000000").NumberFormat = "0.00%"
        Cells(1, lastcol + 5).Value = "Total Stock Volume"
        Cells(2, lastcol + 8).Value = "Greatest % Increase"
        Cells(3, lastcol + 8).Value = "Greatest % Decrease"
        Cells(4, lastcol + 8).Value = "Greatest Total Volume"
        Cells(1, lastcol + 9).Value = "Ticker"
        Cells(1, lastcol + 10).Value = "Value"
        greatestPer = 0
        smallestPer = 0
        greatestTotal = 0
        
        
        'assign value for open cost
        opencost = Cells(2, 3).Value
        totalvol = 0
        
        'start for loop to work on each row
        For i = 2 To lastrow
           
           'if the tickers are different
            If Cells(i + 1, 1).Value <> Cells(i, 1) Then
                
                'get & assign ticker code
                ticker = Cells(i, 1).Value
                Cells(tablerow, lastcol + 2).Value = ticker
                
                'calculate & assign  yearly change
                closecost = Cells(i, lastcol - 1).Value
                yearlychange = closecost - opencost
                Cells(tablerow, lastcol + 3).Value = yearlychange
                    
                    'Assign correct color
                    If Cells(tablerow, lastcol + 3).Value < 0 Then
                        Cells(tablerow, lastcol + 3).Interior.ColorIndex = 3
                    Else
                        Cells(tablerow, lastcol + 3).Interior.ColorIndex = 4
                    End If
                
                'calculate & assign  percent change
                If opencost <> 0 Then
                    percentchange = yearlychange / opencost
                Else
                    percentchange = yearlychange
                End If
                Cells(tablerow, lastcol + 4).Value = percentchange
                
                'calculate & assign  total stock volume
                totalvol = totalvol + Cells(i, lastcol).Value
                Cells(tablerow, lastcol + 5).Value = totalvol
                
                'move on to the next summary table row and reset data
                tablerow = tablerow + 1
                totalvol = 0
                opencost = Cells(i + 1, 3).Value
    
            'if two row have the same ticker
            Else
                totalvol = totalvol + Cells(i, 7).Value

            End If
            
        Next i
        
        'Format to percentage
        Range("Q2:Q3").NumberFormat = "0.00%"
        
        'Look up and store the value of greatest % increase change and its ticker
        greatestPer = Application.Max(Range(Cells(2, lastcol + 4), Cells(lastrow, lastcol + 4)))
        Cells(2, lastcol + 10).Value = greatestPer
        For j = 2 To lastrow
            If Cells(j, lastcol + 4).Value = greatestPer Then
                Exit For
            End If
        Next
        Cells(2, lastcol + 9).Value = Cells(j, lastcol + 2).Value
        
        'Look up and store the value of greatest % decrease change and its ticker
        smallestPer = Application.Min(Range(Cells(2, lastcol + 4), Cells(lastrow, lastcol + 4)))
        Cells(3, lastcol + 10).Value = smallestPer
        For k = 2 To lastrow
            If Cells(k, lastcol + 4).Value = smallestPer Then
                Exit For
            End If
        Next
        Cells(3, lastcol + 9).Value = Cells(k, lastcol + 2).Value

        'Look up and store the value of greatest total volume and its ticker
        greatestTotal = Application.Max(Range(Cells(2, lastcol + 5), Cells(lastrow, lastcol + 5)))
        Cells(4, lastcol + 10).Value = greatestTotal
        For l = 2 To lastrow
            If Cells(l, lastcol + 5).Value = greatestTotal Then
                Exit For
            End If
        Next
        Cells(4, lastcol + 9).Value = Cells(l, lastcol + 2).Value
        
        'Format all cells to fit the value
        Range(Cells(1, lastcol + 2), Cells(lastrow, lastcol + 10)).Columns.AutoFit
        
    Next ws

End Sub


