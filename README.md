# stock-analysis
Performing an analysis on stock data to uncover trends and to create a refactored macro tool through VBA for the client to use to continue to complete analysis on their own.

## Overview of Project
* Using stock data provided to complete an analysis on specific stocks and to create a refactored macro/tool for the client.

### Purpose
* The purpose of this analysis was to indentify trends of specific stocks for years of 2017 & 2018, these stocks were indentified by the client. The analysis was used to help assess and plan for their future investments.
* The client needed a tool to be able to quickly and effectively assess stocks to indentify trends on their own.

## Analysis
### Analysis of stocks for the year of 2017.
* Overall for the year 2017, stocks were trending positively across the board.
* Tickers "DQ" and "SEDG" had the largest gain in volume.
* Only Ticker "TERP" had a negative performance.

![goals](VBA_Challenge_2017.PNG)

### Analysis of stocks for the year of 2018.
* 2018 overall had a different story with a majority of the stocks trending down.
* However, tickers "RUN" and "ENPH" are still trending upwards making them very attractive.
* As for execution time, both analysis for year of 2017 and 2018 took the same amount of time.
* It's also important to note that the refactored code ran faster than the original code.

![goals](VBA_Challenge_2018.PNG)

## Summary
* Refactored code by restructing existing code without changing the behavior of the code.
* A specific example of refactored code in this VBA script was to a "tickerIndex" variable.
  - In creating this "tickerIndex" variable it allowed the macro to run faster because it was was accessing the correct index in four different arrays.
  - **Orginal Loop Code:
     -     For i = 0 to 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
       Worksheets("2018").Activate
       For j = 2 to RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then
               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
               startingPrice = Cells(j, 6).Value
           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
               endingPrice = Cells(j, 6).Value
   - **Refactored Loop Code:
     -     ' Create a ticker Index
    tickerIndex = 0

    ' Create three output arrays for "volume", "starting price" and "ending price"
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ' Create a for loop to initialize the tickerVolumes to zero.
     For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
    
    ' Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
   
        ' Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        ' Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
               tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        ' check if the current row is the last row with the selected ticker
        '' If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
               tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
               tickerIndex = tickerIndex + 1
        End If

            ' Increase the tickerIndex.
        If Cells(i, 1).Value = tickerIndex Then
               tickerIndex = tickerIndex + Cells(i, 6).Value
        End If
### Advantages of Refactored
* Refactoring removes "Code Smell".
* Makes the code easier to maintain.
* Reduces code size to perform faster.

### Disadvantages of Refactored
* May introduce new bugs that are difficult to troubleshoot.
* Can potentially take more time refactor code.
