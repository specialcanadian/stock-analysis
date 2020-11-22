# VBA Stock Analysis

## Overview of Project

Steve requested A VBA program he could use to study DQ’s stock performance during 2017 and 2018. After using the program to analyze DQ stock, Steve wanted to do a similar analysis on the rest of the stock market. To handle the much larger amount of data, I refactored the code with the goal of having it run more efficiently.  

## Results

### Code Analysis
The code was refactored following the specific Canvas instructions and runs much faster than before. Here are the current time results for the refactored code, and the original (Fig.1, Fig.2).

## Fig.1: Refactored code runtime for 2017 and 2018 
![alt text](https://github.com/specialcanadian/stock-analysis/blob/main/Resources/GithubImgRefactoredCode.png?raw=true)

## Fig.2: Original code runtime for 2017 and 2018
![alt text](https://github.com/specialcanadian/stock-analysis/blob/main/Resources/GitHubImgOldCode.png?raw=true)  


  As we can see, the refactored code is about 0.4s faster than the original.
The refactored code’s superior performance comes from removing a nested for loop.   

```
"Original Code"
For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
       Worksheets(yearValue).Activate
       For j = 2 To RowCount
```
This nested loop made the macro run through all the rows of the dataset 11 times, restarting each time the ticker changed.

```
   "Refactored code loop"
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
    
    '2b) Loop over all the rows in the spreadsheet.
    For j = 2 To RowCount
```
In the refactored code, the loops are separate, so that the 2nd loop only need to go through the data set once.

### Stock Analysis
After running the code once for each year, we have two tables (Table.1 and Table.2) showing the performance of multiple stocks. 
## Table.1: Performance of stocks in 2017
![alt text](https://github.com/specialcanadian/stock-analysis/blob/main/Resources/Stock_performance_2017.png?raw=true)

## Table.2: Performance of stocks in 2018
![alt text](https://github.com/specialcanadian/stock-analysis/blob/main/Resources/Stock_performance_2018.png?raw=true)
## Summary

###


