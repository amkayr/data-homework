Sub StockMarketAnalysis_2016():
   Dim ticker As String
   Dim totalVolume As String
   totalVolume = 0
   Dim lastRow As Long
   lastRow = Cells(Rows.Count, 1).End(xlUp).Row
   
   Dim yearlyChange As Variant
   yearlyChange = 0
   Dim openYear As Double
   Dim closeYear As Double
   Dim percentChange As Variant
   
   Dim i As Long

   Dim summaryTableRow As Integer
   summaryTableRow = 2
   
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"



   For i = 2 To lastRow

       
       currentValue = Cells(i, 1).Value
       nextValue = Cells(i + 1, 1).Value
     
     

       If nextValue <> currentValue Then
           ticker = currentValue
           totalVolume = totalVolume + Cells(i, 7).Value
        
        
           openYear = Range("C" & i - 261).Value
           closeYear = Range("F" & i).Value
           
           If openYear <> 0 Then
           
                yearlyChange = CDec(closeYear - openYear)
                percentChange = (closeYear / openYear) - 1
                
                Else
                
                yearlyChange = CDec(closeYear - openYear)
                percentChange = 0
                
                End If
                
                     

           
           Range("I" & summaryTableRow).Value = ticker
           Range("L" & summaryTableRow).Value = totalVolume
           
                
           Range("J" & summaryTableRow).Value = CDec(yearlyChange)
                If yearlyChange <= 0 Then
                Range("J" & summaryTableRow).Interior.ColorIndex = 3
                Else
                Range("J" & summaryTableRow).Interior.ColorIndex = 4
                
                End If
                
        
            Range("K" & summaryTableRow).Value = Format(percentChange, "Percent")


           summaryTableRow = summaryTableRow + 1
           totalVolume = 0
         

       Else
            totalVolume = totalVolume + Cells(i, 7).Value
            openYear = Range("C" & i).Value
            closeYear = Range("F" & i).Value
            
       
        


       End If


      Next i


End Sub