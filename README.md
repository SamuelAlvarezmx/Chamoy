# Stock Analysis
## Overview of Project

We will use VBA within excell to analyze the stock market for green stocks and see how they have performed in 2017 and 2018. In order to make this happen it was evaluated the return percentage and the Total Daily Volume to see how well each of the stocks performed. We will need to refactor a code used before in order to enhance the time process and make it faster than the original one in order to make it fitter in any case that the data might grow larger.

## Results
### Code

The following is the code refactored as used in my VBA_ challenge workbook.
"Sub AllStocksAnalysisRefactored()

    'Set the timer for the process
    Dim startTime As Single
    Dim endTime As Single
    
    'Input box to determine the year the analysis will be performed
    
    yearValue = InputBox("What year would you like to run the analysis on?")
    
    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    
    Worksheets("All Stocks Analysis").Activate
    
    'Create a header row
    
    Cells(1, 1).Value = "All Stocks (" + yearValue + ")"
    
     Cells(1, 1).Font.Bold = True
     Cells(1, 1).Font.Size = 16
     
    Cells(3, 1).Value = "Ticker"
    
    Cells(3, 2).Value = "Total Daily Volume"
    
    Cells(3, 3).Value = "Return"
    
    Range("A1").Font.ColorIndex = 15
    
    Range("A3:C3").Font.Bold = True
    
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    Range("A3:C3").Font.Size = 14
    
    Range("A3:C3").Interior.ColorIndex = 14
    'Initialize array of all tickers

    Dim Tickers(12) As String

    Tickers(0) = "AY"
    Tickers(1) = "CSIQ"
    Tickers(2) = "DQ"
    Tickers(3) = "ENPH"
    Tickers(4) = "FSLR"
    Tickers(5) = "HASI"
    Tickers(6) = "JKS"
    Tickers(7) = "RUN"
    Tickers(8) = "SEDG"
    Tickers(9) = "SPWR"
    Tickers(10) = "TERP"
    Tickers(11) = "VSLR"
    
    'Activate Data worksheet
     
    Worksheets(yearValue).Activate
    
    'Get the number of rowsa to loop over
    
    rowstart = 2
    RowEnd = Cells(Rows.Count, 1).End(xlUp).Row
    
    '1a Create a ticker index
    
    Dim tickerIndex As Integer
    
    tickerIndex = 0
    
    '1b Create three output arrays
    
    Dim tickerVolumes(12) As Long
    
    Dim tickerStartingPrices(12) As Single
    
    Dim tickerEndingPrices(12) As Single
    
    Dim i As Integer
    
    
   '2a Create a for loop to initialize the tickerVolumes to zero.
    
        
    For i = 0 To 11
        
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
        
    Next i
        
        '2b Loop over all the rows in the spreadsheet
    
        For i = rowstart To RowEnd

            ' 3a Increase totalVolume of ticker
        
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
           
            '3b Check if the current row is the first row with the selected tickerindex
        
            If Cells(i, 1).Value = Tickers(tickerIndex) And Cells(i - 1, 1).Value <> Tickers(tickerIndex) Then
        
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
            End If
        
            '3C check if the current row is the lastrow with the selected ticker
        
            If Cells(i, 1).Value = Tickers(tickerIndex) And Cells(i + 1, 1).Value <> Tickers(tickerIndex) Then
        
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                    
            tickerIndex = tickerIndex + 1
        
            End If
    
        Next i
        
        For i = 0 To 11
            
            tickerIndex = i
            
        
        '4 create a for loopto loop through the arrays to output results
        
        Worksheets("All Stocks Analysis").Activate
    
        Cells(4 + i, 1).Value = Tickers(tickerIndex)
        Cells(4 + i, 2).Value = tickerVolumes(tickerIndex)
        Cells(4 + i, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
        
        Next i
        
        'Change format of the data
        
        For i = 0 To RowEnd
            
            Cells(4 + i, 2).NumberFormat = "$#,##.00"
    
            Cells(4 + i, 3).NumberFormat = "0.0%"
    
    
            If Cells(4 + i, 3) < 0 Then
    
                Cells(4 + i, 3).Interior.Color = vbRed
    
            ElseIf Cells(4 + i, 3) > 0 Then
    
                Cells(4 + i, 3).Interior.Color = vbGreen
    
            Else
    
                Cells(4 + i, 3).Interior.Color = xlNone
    
            End If
    
        Next i
    
    endTime = Timer
    
    'MsgbOX that will tell the time the subroutine ran
    
    MsgBox ("This code ran in" + " " + Str(endTime - startTime) + " " + "seconds for the year " + (yearValue))
    
End Sub"

### Comparison between 2018 and 2017

Furing 2017, almost all green stocks had a positive return. The most relevant was "DQ" with almost a 200% return during this year. However, stepping into 2018 the green stocks plumbetted to almost all of them having negative returns. In this case the only ones that had a positive year were "RUN" and "ENPH".


## Summary

Refactoring code is a common practice because there are sometimes that code can be more easier and therefore faster, there is no need to go through the same amount of loops or some data changed. The advantages of refactoring is that you have the felxibility to get the results that you are looking for. More efficiency or faster procedures. Disadvantages can be also that understanding the code of someone else is not very clear and maybe some lines of code that you think are not essential for the whole subroutine may not be true in the end. Also, refactoring can change completely the purpose of the first one.
In this particular case, I think the main cons is that we are using more variables, is not as straight forward as the first subroutine made. If you are not following correctly your steps it might get confusing while refactoring.

The pros is that with all these new variables and arrays I think the subroutine has more flexibility, even if the data gets bigger. Also it is important organizing the information in arrays, letting the program run faster through the lines of code. It is seen in the comparison of the time it takes to run the code between the refactored and the original one.
### Time running code before refactoring
<img width="742" alt="2017-code" src="https://user-images.githubusercontent.com/104656920/178124618-38a9f404-936c-4f15-9614-443a6fa9af91.png">
### Time running code after refactoring
<img width="741" alt="VBA_Chanllenge_2017" src="https://user-images.githubusercontent.com/104656920/178124622-d6b4b502-0988-4f8c-8e9a-3adf0b71d72e.png">
