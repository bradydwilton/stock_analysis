# VBA Challenge - Refactor Stock Analysis Subroutine
## Overview of Project

### Purpose
This project began as an attempt to use VBA to analyze stock data for twelve green energy companies. A VBA script was written to analyze the data and format the output, however, the code took over 1 second to run for the twelve companies. The purpose of this project is to refactor the script to be able to run much more efficiently, allowing the subroutine to scale and handle data for any number of companies.  

## Results

### Performance Prior to Refactoring  
Initially, the code was designed to loop through through each of the twelve tickers (held in an array). During each loop through, three nested loops pass through each row of data to calculate the total volume and yearly return for each ticker.    
  
##### For loop through each of the twelve tickers
>     'Create nested loops to find total volume and yearly return  
>     For i = 0 to 11  
>         'loop to calculate total trading volume  
>         'loop to calculate initial stock price  
>         'loop to calculate final stock price  
>     Next i  
  
##### Nested For loop used to calculate trading volume during each loop for individual tickers  
>         'Calculate total daily volume
>         For j = 2 To rowEnd
>            If Cells(j, 1).Value = ticker Then
>                totalVolume = totalVolume + Cells(j, 8).Value
>            End If
>         Next j  

Similar nested For loops were used to find the starting and ending price of each ticker, which then could be used to calculate the yearly return. Each loop executed over 3,013 rows of data once for each ticker. This relatively large computation was the cause for the run times between 1.2 and 1.4 seconds, as shown below.

<img src="https://github.com/bradydwilton/stock_analysis/blob/main/resources/first_draft_2017_run.png" width=900>  
_**The image above shows the time of the run with the 2017 data before refactoring**_

<img src="https://github.com/bradydwilton/stock_analysis/blob/main/resources/first_draft_2018_run.png" width=900>
_**The image above shows the time of the run with the 2018 data before refactoring**_

### Areas of Improvement

### Performance After Refactoring

## Summary

1. Advantages and disadvantages of refactoring code

2. How do these pros and cons apply to refactoring the original VBA script?
