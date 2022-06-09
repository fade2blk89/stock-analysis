# 2017-2018 Stock Return Analyses #
## Overview of Project ##
The purpose of this project is to analyze a large cross-section of stock ticker and volume data for various companies to help determine their returns in the years 2017 and 2018. 

## Results ##

When analyzing the data, I found that in 2017, the company Daqo New Energy Corp (DQ) had a return of almost 200%, as shown below:
![All Stocks 2017](https://github.com/fade2blk89/stock-analysis/blob/main/All%20Stocks%202017.png)

Conversely, in 2018, it appears that a vast majority of the companies had a negative return on investment, with only Enphase Energy Inc (ENPH) and Sunrun Inc (RUN) reporting a positive return for both years:

![All Stocks 2018](https://github.com/fade2blk89/stock-analysis/blob/main/All%20Stocks%202018.png)

A key part of refactoring the code involved quickly verifying that the desired company stock ticker matched our current index in the array. If it did not match, it would check both the previous line and following line before moving onto the next index in the array. This was accomplished with the following: 

    If Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
            
      tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
    
    End If
                
    If Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
            
      tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
    
      tickerIndex = tickerIndex + 1
     
    End If

## Summary ##

Based on our findings, it appears that Enphase and Sunrun have experienced the best stock return performance. After refactoring the code using the above conditional formatting, I was able to run the 2017 and 2018 analyses in less than a fifth of a second, as shown below: 

![VBA Challenge 2017](https://github.com/fade2blk89/stock-analysis/blob/main/VBA_Challenge_2017.png)

![VBA Challenge 2018](https://github.com/fade2blk89/stock-analysis/blob/main/VBA_Challenge_2018.png)

In general, refactoring code and clearly defining the parameters using tools such as nested loops and conditional formatting allows us to simplify the code being written and ideally will have the macro run more quickly and smoothly, but it may also result in errors if not refactored correctly. One advantage of the refactored VBA script is that it not only shortened the result time, but was also better formatted with percentages and color coding. 
