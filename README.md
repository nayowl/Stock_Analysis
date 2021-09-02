# Refactoring Code for Stock Analysis

## 1 Overview for The Project

### 1.1 Background

Stock analysis is a method evaluation of stocks for investor to make a buying or selling decision. The current analysis is using VBA excel to analyze green stock data ( green_stock.xlsm) in 2017 and 2018. It provides Total Daily Volume and Return of each stock in selected year. Return calculated by comparing the difference between starting and ending price to the starting price. 

The existing code works well for calculate the current data since the data size is moderate. However,  the code needs improvement to calculate larger data size.  Refactoring is a disciplined technique for improving and restructuring the design of an existing code base without changing its external behavior. This project will refactor the existing code to improve the performance of the code. 

### 1.2 Purpose

-	To refactoring the existing stock analysis code, therefore the code will be more efficient and maintainable. 

-	Comparing the existing code to the refactoring code to find the difference result performance. 

## 2 Results

The original dataset in the excel contains of two sheet which is 2017 and 2018. Each sheet holds the stock data of the year. Table 1 below describe the summary of data used for analysis. It has 12 Ticker with 251 row data on each ticker. On the dataset there are 8 columns which are ticker, date, open, high, low, close, adj close and volume. Table 2 below will show us the mapping of data column and its usage in analysis, if the column is not present in the table, then it is not used in analysis.  

<p align="center">
<sub>Table 1 Summary of stock Analysis Data</sub>
</p> 

|Stock Name(Ticker Name) 2017|Data Count 2017|Stock Name(Ticker Name) 2018|Data Count 2018|
|---|---|---|---|
|AY|251|AY|251|
|CSIQ|251|CSIQ|251|
|DQ|251|DQ|251|
|ENPH|251|ENPH|251|
|FSLR|251|FSLR|251|
|HASI|251|HASI|251|
|JKS|251|JKS|251|
|RUN|251|RUN|251|
|SEDG|251|SEDG|251|
|SPWR|251|SPWR|251|
|TERP|251|TERP|251|
|VSLR|251|VSLR|251|
|Grand Total|3012|Grand Total|3012|

<p align="center">
<sub>Table 2 Column Usage in Analysis</sub>
</p>

|Column Name|Usage in Analysis|
|---|---|
|Ticker|Ticker Name|
|Close|* Starting price if it’s the first data row of each ticker, * Ending Price if it’s the last data row of each ticker|
|Volume|To calculate total Volume|

The Subroutine AllStocksAnalysis() will be used to calculate and put the data result in excel sheet. The main difference between the existing code and the refactoring code is the usage of arrays to hold volume, starting price and ending price value in the refactoring code instead of nested loop to find the value.

In existing code, to calculate volume and define starting and ending price, we used nested loops. 
```
For i = 0 To 11
            ticker = tickers(i)
            totalVolume = 0
            Worksheets(yearValue).Activate
            
        For j = 2 To RowCount
        ' increase totalVolume if the ticker value (row A) is DQ
        
            'To count totalvolume
            If Cells(j, 1) = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
        
            'To count starting price for Yearly Return
            If Cells(j - 1, 1) <> ticker And Cells(j, 1) = ticker Then
                startingPrice = Cells(j, 6).Value
            End If
        
            'To count Ending price for Yearly Return
            If Cells(j + 1, 1) <> ticker And Cells(j, 1) = ticker Then
                EndingPrice = Cells(j, 6).Value
            End If
            
        Next j
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = EndingPrice / startingPrice - 1
        
    Next i

```
The First loop ( For i =0 to 11)  is used to loop through the tickers. Tickers is an array that hold all the ticker name. The second loop (nested from the fist loop) (for j=2 to RowCount) is used to loop through rows in the data to find total volume, starting price, and ending price of the ticker from the first loop.  

For the refactoring code, arrays will be used to hold total volume, starting price and ending price of each ticker using ticker index.  There are no nested loops in the refactoring code. First the code creates a loop to initialize ticker volume to zero, then loop through the row using TickerIndex to access the correct index of each ticker to count total volume and define starting and ending price. The ticker index value will be added when the next row ticker doesn’t match with the current ticker

```
'1a) Create a ticker Index
     
TickerIndex = 0
    
    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
     For j = 0 To 11
            tickerVolumes(j) = 0
     Next j
            
    ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
    
            '3a) Increase volume for current ticker
        
                tickerVolumes(TickerIndex) = tickerVolumes(TickerIndex) + Cells(i, 8).Value
           
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
                If Cells(i - 1, 1) <> tickers(TickerIndex) Then
                    tickerStartingPrices(TickerIndex) = Cells(i, 6).Value
        'End If
                End If
            
    
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
                If Cells(i + 1, 1) <> tickers(TickerIndex) Then
                    tickerEndingPrices(TickerIndex) = Cells(i, 6).Value
        
            '3d Increase the tickerIndex.
                    TickerIndex = TickerIndex + 1
        'End If
                End If
        Next i
            
            
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        For i = 0 To 11
        
            Worksheets("All Stocks Analysis").Activate
            Cells(4 + i, 1).Value = tickers(i)
            Cells(4 + i, 2).Value = tickerVolumes(i)
            Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        Next i

```
The performance of the code will be measured by how long it takes to execute the code. Figure 1  and Figure 2 will show us times to execute existing code  for 2017 and 2018 data. While Figure 3 and Figure 4 show execution times for the code after refactoring. For comparison between existing and refactoring code of the execution time we can see it in the Table 3. 


<img width="621" alt="green_stock_2017" src="https://user-images.githubusercontent.com/88597187/131785286-4bc66fc4-dc42-4144-9f75-9d381a6824c0.png">

<sub>Figure 1 Execution times for 2017 data in Existing Code</sub>
 
<p>&nbsp;</p>
<p>&nbsp;</p>


<img width="620" alt="green_stock_2018" src="https://user-images.githubusercontent.com/88597187/131785300-53e1c0a3-0e48-43a3-a5ea-f64961952580.png">

<sub>Figure 2 Execution times for 2018 data in Existing Code</sub>

<p>&nbsp;</p>
<p>&nbsp;</p>


<img width="633" alt="vba_challenge_2017" src="https://user-images.githubusercontent.com/88597187/131785312-f061c89d-d521-4efa-bef2-f366b2b04af1.png">

<sub>Figure 3 Execution times for 2017 data in Refactoring Code</sub>

<p>&nbsp;</p>
<p>&nbsp;</p>



<img width="632" alt="vba_challenge_2018" src="https://user-images.githubusercontent.com/88597187/131785319-d43f7216-3959-4fae-a5e4-619c65f1a42b.png">

<sub>Figure 4 Execution times for 2018 data in Refactoring Code</sub>

<p>&nbsp;</p>
<p>&nbsp;</p>



<p align="center">
<sub>Table 3 Comparison of Execution Times </sub>
</p>

|Year|Existing|Refactoring|Difference (Existing-Refactoring)|Reduction in time(Existing- Refactoring)/Existing x100%|Increase in Performance(Existing-Refactoring)/Refactoring x100%|Time multiplication (Existing/Refactoring)|
|---|---|---|---|---|---|---|
|2017|2.836|0.375|2.461|86.777|656.267|7.563|
|2018|2.930|0.195|2.735|93.345|1402.564|15.026|

*The values show in this table are roundup to 3 decimals*

From Table 3 we can see the positive improvement for the code after refactoring. The reduction in time for executing the code is 86.777% and 93.345%. Furthermore, the increase in performance and the multiplication shows us the code is become more efficient. There are 656.267% and 1402.564% increase in performance and the code become 7.563 and 15.026 faster than before the refactoring happened. 


## 3 Summary

### 3.1 Advantages and Disadvantages of Refactoring Code
Advantages:
- The code is easier to enhance and maintain in the future
- Less complex and easier to read
- Prevents many future defects
- May increase performance

Disadvantages:
- It may introduce bugs
- Takes times and expensive in budget
- Risky if the application is too big and there is not proper test case in existing code
- Risky if the developer does not understand what’s all about

### 3.2 Pros and Cons Refactoring Code in this Project

From the result above, it’s clear that refactoring has positive impact on the existing code in this project. Its increase the performance and make the execution time faster. The factor that may lead the improvement is the usage of array instead of nested loops for calculation. For the disadvantage of the refactoring, since we only have two dataset (2017 and 2018) to compare  and doesn’t have another test case it may introduce bugs and will not work as expected.
