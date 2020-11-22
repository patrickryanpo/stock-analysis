# stock-analysis

## Overview of Project

The purpose of this project is to use VBA to represent green stock data for years 2017 and 2018. This project was created for Steve in order to aid him and his parents in looking at yearly data trends, which will allow them to make informative future investments. 

The code in this project has been refactored to run efficiently on more systems, and will allow analysis to run easily through additional stock data. 

## Results 

By refactoring the code, it can be seen in the photos below that the time it takes to run the code significantly dropped with the refactored code when compared to the original: 

*2017 Refactored Run Time:*

![Image of 2017 Refactored Run Time](https://github.com/patrickryanpo/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

<br />

*2017 Original Code Run Time:*

![Image of 2017 Original Run Time](https://github.com/patrickryanpo/stock-analysis/blob/main/Resources/2017%20Original%20Code.png)

As we can see from the representation above, the refactored code run in 0.08 seconds when compared to the original code with run in 0.4 seconds. The same conclusion can be made with the 2018 analysis, illustrated below. 

*2018 Refactored Run Time:*

![Image of 2018 Refactored Run Time](https://github.com/patrickryanpo/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

*2018 Original Code Run Time:* 

![Image of 2018 Original Code Run Time](https://github.com/patrickryanpo/stock-analysis/blob/main/Resources/2018%20Original%20Code%20Run%20Time.png)

When it comes to the code, the nesting order of the loops were changed. In order for this to happen, an additional array was created that was named tickerVolumes as compared to the previous code which only had the two other arrays (startingPrice and endingPrice).

Here is the original code: 

```
Sub AllStocksAnalysis()

Dim startTime As Single
Dim endTime As Single

yearValue = InputBox("What year would you like to run the analysis on?")
    startTime = Timer
    '1.  Format the output sheet on the "All Stocks Analysis" worksheet.
    Worksheets("All Stocks Analysis").Activate
         
        Range("A1") = "All Stocks (" + yearValue + ")"
    
        'Header Row
        Cells(3, 1) = "Ticker"
        Cells(3, 2) = "Total Daily Volume"
        Cells(3, 3) = "Return"
        
    '2.  Initialize an array of all tickers.
    
    Dim tickers(12) As String
    
        tickers(0) = "AY"
        tickers(1) = "CSIQ"
        tickers(2) = "DQ"
        tickers(3) = "ENPH"
        tickers(4) = "FSLR"
        tickers(5) = "HASI"
        tickers(6) = "JKS"
        tickers(7) = "RUN"
        tickers(8) = "SEDG"
        tickers(9) = "SPWR"
        tickers(10) = "TERP"
        tickers(11) = "VSLR"
        
    '3.  Prepare for the analysis of tickers.
    
        '3a   Initialize variables for the starting price and ending price.
        
        Dim startingPrice As Single
        Dim endingPrice As Single
            
        '3b   Activate the data worksheet.
        
        Worksheets(yearValue).Activate
            
        '3c  Find the number of rows to loop over.
        
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
            
    '4.  Loop through the tickers.
    
        For i = 0 To 11
        
            ticker = tickers(i)
            totalVolume = 0
   
    '5.  Loop through rows in the data.
    
        Worksheets(yearValue).Activate
            For j = 2 To RowCount
    
        '5a   Find the total volume for the current ticker.
        
        If Cells(j, 1).Value = ticker Then
            totalVolume = totalVolume + Cells(j, 8).Value
            
        End If
        '5b   Find the starting price for the current ticker.
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        
            startingPrice = Cells(j, 6).Value
        
        End If
        
        '5c   Find the ending price for the current ticker.
         If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
            endingPrice = Cells(j, 6).Value
        
        End If
        
        Next j
        
    '6.  Output the data for the current ticker.
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
        
        Next i

endTime = Timer
MsgBox "This code ran in" & (endTime - startTime) & "seconds for the year" & (yearValue)

End Sub

```

Here is the refactored code:

```
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    
    tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    
        For j = 2 To RowCount
    
    
        '3a) Increase volume for current ticker
        
             tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
        If Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
        End If
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
    
         If Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(j, 6).Value

            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next j
    
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
For i = 0 To 11
        
    Worksheets("All Stocks Analysis").Activate
        
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
```
## Summary 

As presented in the results section of this Readme file, refactoring code allows for a more efficient way of running the code. In line with this, refactoring code is important in improving the code architecture, which will allow the code to be more understandable, run faster, and consistent. Refactored code also plays a key role in reducing technical costs that allows a cleaner code that will be valuable in the future. 

Disadvantages of refactoring code include time consumption and road blocks. At times, refactoring code may not be as straight forward and may take more time than projected. The person refactoring the code may also hit roadblocks which may hinder the project from moving forward. 

The same advantages and disadvantages could be said about VBA scripts. In this example, it took me longer to create the refactored code as compared to writing the original code. The time to debug the refactored code took longer than the time saved running with the refactored code. In a small data set, it is not ideal to refactor code as it will only save you a fraction of a second. However, when dealing with bigger data sets, I believe refactoring will prove to be more worthwhile. 

With this, I believe that it is best practice to always save the original code prior to refactoring. 
