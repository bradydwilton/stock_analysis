# VBA Challenge - Refactor Stock Analysis Subroutine
## Overview of Project

### Purpose
This project began as an attempt to use VBA to analyze stock data for twelve green energy companies. A VBA script was written to analyze the data and format the output, however, the code took over 1 second to run for the twelve companies. The new scope of the project is to refactor the script to be able to run much more efficiently, allowing the subroutine to scale and handle data for any number of companies.  

## Results

### Performance Prior to Refactoring  
Initially, the code was designed to loop through through each of the twelve tickers (stored in an array). During each loop through, three nested loops pass through each row of data to calculate the total volume and yearly return for each ticker.    
  
##### For loop through each of the twelve tickers

``` vb
     'Create nested loops to find total volume and yearly return  
     For i = 0 to 11
         
         'loop to calculate total trading volume  
         
         'loop to calculate initial stock price  
         
         'loop to calculate final stock price  
         
     Next i  
```

##### Nested For loop used to calculate trading volume during each loop for individual tickers  

``` vb
         'Calculate total daily volume
         For j = 2 To rowEnd
         
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
         
         Next j  
```

Similar nested For loops were used to find the starting and ending price of each ticker, which then could be used to calculate the yearly return. Each loop executed over 3,013 rows of data once for each ticker. This relatively large computation was the cause for the run times between 1.2 and 1.4 seconds, as shown below.

<img src="https://github.com/bradydwilton/stock_analysis/blob/main/resources/first_draft_2017_run.png" width=900>

_**The image above shows the subroutine's run time with the 2017 data before being refactored**_  

<img src="https://github.com/bradydwilton/stock_analysis/blob/main/resources/first_draft_2018_run.png" width=900>

_**The image above shows the subroutine's run time with the 2018 data before being refactored**_

### Areas of Improvement

As previously mentioned, the original design of the subroutine consisted of a parent loop, which looped through each ticker, and three nested loops, which performed the calculations presented in the analysis. To reduce the number of computations done by the subroutine (and thus reduce the run time), the sub needed to be refactored to loop through each row, performing all necessary calculations on the individual row at once.  
  
To reduce the runtime of the sub, the following loop, along with three arrays to hold the data and a ticker index to determine the current ticker and assign data to the appropriate elements of each array, were used in the refactored subroutine.    
  
##### The final loop is shown below with portions of the body removed for the purpose of this report:

``` vb
     For i = 2 to rowCount
    
        'Increase volume for current ticker
        If Cells(i, 1) = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8)
        End If
    
        'If statement to determine if row is the current ticker's starting price   
        'and assign to tickerStartingPrices(tickerIndex) if so
        
        'If statement to determine if row is the current ticker's ending price  
        'and assign to tickerEndingPrices(tickerIndex) if so
        
        'Increment the ticker index
        tickerIndex = tickerIndex + 1
    
     Next i  
```


### Performance After Refactoring  
  
After refactoring the subroutine as described above, the runtimes were reduced from the original range of 1.2 - 1.4 seconds to _**under 0.25 seconds!**_ 

<img src="https://github.com/bradydwilton/stock_analysis/blob/main/resources/refactored_2017_run.png" width=900>

_**The image above shows the subroutine's run time with the 2017 data after being refactored**_  

<img src="https://github.com/bradydwilton/stock_analysis/blob/main/resources/refactored_2018_run.png" width=900>

_**The image above shows the subroutine's run time with the 2018 data after being refactored**_  

The refactored subroutine is now over _**80% more efficient**_ than the original sub and ready to be scaled to analyze a much larger number of stock data.

## Summary

##### Advantages and disadvantages of refactoring code

The advantages of refactoring code are clear - the initial solution to a problem solved by code is likely not the most efficient solution! Refactoring code takes extra time, but is necessary to create clean, sharable, and scalable solutions.

##### How do these pros and cons apply to refactoring the original VBA script?

The origianal VBA script performed the necessary computations to analyze the given data, creating a possible argument that it was fine as-is and the time to refactor was not needed. After refactoring, however, the code saw a greater than 80% improvement in run time and is now scalable to handle much more data, showing how important it is to take the time to refactor a solution before accepting the final results.
