Attribute VB_Name = "Module3"
Sub WorksheetLoop()

    ' Define and Initialize variables
    
    
    Dim ws As Worksheet
    
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryrow As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseStock As String
    Dim greatestDecreaseStock As String
    Dim greatestVolumeStock As String
    
    
    
    
 
    
    For Each ws In Worksheets
    
    ws.Activate
    
    
    
    
    
    ' Initialize variables
    
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    ticker = Cells(2, 1).Value
    openingPrice = Cells(2, 3).Value
    totalVolume = 0
    summaryrow = 2
    
    
    
    'Add headers for Columns for Each Sheet
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Incease"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    ' Loop through data and compute information
    For i = 2 To lastRow
        
        ' Check if ticker symbol has changed
        
        If Cells(i, 1).Value <> ticker Then
            
            ' Compute yearly change and percent change
            
            yearlyChange = closingPrice - openingPrice
            If openingPrice <> 0 Then
                percentChange = yearlyChange / openingPrice
            Else
                percentChange = 0
            End If
            
            ' Output data to sheet
            
            Range("I" & summaryrow).Value = ticker
            Range("J" & summaryrow).Value = yearlyChange
            Range("K" & summaryrow).Value = percentChange
            Range("L" & summaryrow).Value = totalVolume
            
            
            
           
            
            ' Reset variables for new ticker symbol
            
            ticker = Cells(i, 1).Value
            openingPrice = Cells(i, 3).Value
            totalVolume = 0
            summaryrow = summaryrow + 1
            
            
            
        End If
        
        ' Compute total volume for ticker symbol
        
        totalVolume = totalVolume + Cells(i, 7).Value
        
        ' Set closing price for ticker symbol
        
        closingPrice = Cells(i, 6).Value
        
        
        
        If percentChange > greatestIncrease Then
    
            greatestIncrease = percentChange
            greatestIncreaseStock = Range("I" & (summaryrow - 1)).Value
        
        
        End If
        
        
        If percentChange < greatestDecrease Then
        
            greatestDecrease = percentChange
            greatestDecreaseStock = Range("I" & (summaryrow - 1)).Value
        
        End If
        
        
        If totalVolume > greatestVolume Then
        
            greatestVolume = totalVolume
            greatestVolumeStock = Range("I" & summaryrow).Value
            
        End If
        
        
        
    Next i
    
    
    
    ' Compute information for last ticker symbol
    
    
    yearlyChange = closingPrice - openingPrice
    
    
    
    If openingPrice <> 0 Then
        percentChange = yearlyChange / openingPrice
    Else
        percentChange = 0
        
    
     
        
        
    End If
    
    
    Cells(i, 9).Value = ticker
    Cells(i, 10).Value = yearlyChange
    Cells(i, 11).Value = percentChange
    Cells(i, 12).Value = totalVolume
    Range("P2").Value = greatestIncreaseStock
    Range("Q2").Value = greatestIncrease
    Range("P3").Value = greatestDecreaseStock
    Range("Q3").Value = greatestDecrease
    Range("P4").Value = greatestVolumeStock
    Range("Q4").Value = greatestVolume
    
' Loop through next worksheet
    
Next
    
    
 
    
End Sub




