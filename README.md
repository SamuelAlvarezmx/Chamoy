# Stock Analysis
## Overview of Project

We will use VBA within excell to analyze the stock market for green stocks and see how they have performed in 2017 and 2018. In order to make this happen it was evaluated the return percentage and the Total Daily Volume to see how well each of the stocks performed. We will need to refactor a code used before in order to enhance the time process and make it faster than the original one in order to make it fit to any case for example data growing larger.

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

With these changes in the code, the execution time was quicker in the refactored code than the previous one. Using the arrays to store data and then retrieve it was faster than doing loops to calculate and output the data.

### Execution time

<img width="742" alt="2017-code" src="https://user-images.githubusercontent.com/104656920/178885812-dcee3971-ce1c-4a15-9a33-183711bf7346.png">
<img width="835" alt="2018-code" src="https://user-images.githubusercontent.com/104656920/178885850-5791954c-5aa3-4aad-b98e-d925d2fdae49.png">

### Execution time refactoring

<img width="741" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/104656920/178885903-8204917d-f90b-4a0a-a2a9-5a4c23bbdb70.png">
<img width="774" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/104656920/178885917-4d267511-9795-4496-8d4e-0d292767b451.png">


### Comparison between 2018 and 2017

During 2017, almost all green stocks had a positive return. The most relevant was "DQ" with almost a 200% return during this year. However, stepping into 2018 the green stocks plumbetted to almost all of them having negative returns. In this case the only ones that had a positive year were "RUN" and "ENPH".
![Return_2017](https://user-images.githubusercontent.com/104656920/178869490-aba9cb54-bc37-4f81-8577-da35651aae3f.png)

![Return_2018](https://user-images.githubusercontent.com/104656920/178869573-527d6a47-2dfc-47c0-a3ac-5c84c9568696.png)


## Summary

Refactoring code is a common practice because there are sometimes that code can be easier and therefore faster, there is no need to go through the same amount of loops or some data might change and then the code needs to be adjusted. The advantages of refactoring are the flexibility to get the results that you are looking for, more efficiency and faster procedures or execution times like in this project.
One of the disadvantage of refactoring could be during the interpretation od the code. Understanding the code of someone else could come across as difficult and maybe some lines of code that you think are not essential for the whole subroutine might actually be. Also, refactoring can change completely the purpose of the original code.
In this particular case, I think the main cons is that we are using more variables, is not as straight forward as the first subroutine made. If you are not following correctly your steps it might get confusing while refactoring.

The pros is that with all these new variables and arrays I think the subroutine has more flexibility, even if the data gets bigger. Also it is important organizing the information in arrays, letting the program run faster through the lines of code. It is seen in the comparison of the time it takes to run the code between the refactored and the original one.
