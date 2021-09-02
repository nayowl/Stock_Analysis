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


<sub>Table 1 Summary of stock Analysis Data</sub>

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

<sub>Table 2 Column Usage in Analysis</sub>

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

## 3 Summary
### 3.1 Advantages and Disadvantages of Refactoring Code
### 3.2 Pros and Cons Refactoring Code in this Project
