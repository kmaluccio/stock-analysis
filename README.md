# stock-analysis project

# Overview of Analysis

In this project, we explore green energy stock performance by analyzing financial data using VBA. With data from the client Steve, we will write one VBA macro to complete an initial analysis on parts of the data. Then, we will refactor the same macro to run slightly faster, being able to analyze larger sets of data for more stocks. With the results of our analysis, we will be able to see how the green energy stocks performed in the years 2017 and 2018 to help Steve's parents determine whether or not they should invest in the stock or how much they want to invest if they decide it is a good option.

Purpose of the analysis.

## Results

To write the refactored code, we had to create arrays for the data so that we could store more than one value in 
variables like the starting price, ending price, and volume for each stock. The main for loop and if-then statement is shown below. Here we start the for loop using the total number of rows which we have found using a line of code. Then, we increase the volume for the current stock, check to see if this is the first or last row to assign start and end prices, and then check to see if we need to increase the ticker index once we have moved onto the next stock.

For i = 2 To RowCount            tickerVolumes(tickerIndex) = Cells(i, 8).Value + tickerVolumes(tickerIndex)                If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then                    tickerStartingPrices(tickerIndex) = Cells(i, 6).Value                    End If        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then                        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value            tickerIndex = tickerIndex + 1                    End If    Next i

This code was efficient and ran quickly through all stocks instead of the original code which ran a separate analysis for every stock.When using the original script to run the analysis on all 12 stocks, the run time for my MacBook was about 0.85 seconds to run 2018 stock analysis and 0.84 seconds to run 2017 stock analysis. The time it took to execute the analysis using the refactored code is shown in the following screenshots.

![2018 run time](https://github.com/kmaluccio/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

![2017 run time](https://github.com/kmaluccio/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

As you can see from above, the refactored script ran significantly faster than the original script. Therefore, if we wanted to analyze thousands of different stocks, instead of only 12, then we would definitely want to run the refactored script. Since the time is less than one second no matter which macro is used for this particular analysis, the original script is good enough to run on the given data.

Based on the stock analysis, there are only two stocks with a positive return in 2018 and these are ENPH and RUN. However, in 2017 every stock had a positive return except one (TERP). The top stocks in 2017 were DQ, SEDG, ENPH, and FSLR with each of their returns being above 100% and RUN stock was a low return below 10%. There are different options or approaches one can take to invest in particular stocks based on this data. It would seem reasonable to invest in ENPH since this stock had a positive return both years and had one of the highest in 2017. Considering other factors may help to give reason to invest in other stocks, but based on this data that seems to be the best option and least risky (although there is always risk in the stock market). 

###VBA macro scripts and worksheets with analyses can be found in the excel file: [VBA Challenge](https://github.com/kmaluccio/stock-analysis/blob/main/VBA_Challenge.xlsm)

## Summary

The first macro written was able to analyze Daqo stock and we were able to see the total volume and return for a given year. Then, a macro was written to analyze all stocks and give this output for each stock in our data set. We were able to highlight those with positive return in green and negative return in red, so that it is easy to see which stocks did well in the given year and which did not. In the excel worksheet, buttons were added to clear the worksheet and perform the analysis of the data which allows for quick results and easy access to run the macro. 

Some challenges in writing these macros were keeping track of the for loops and it-then statements. Initially, when the refactored macro was run for the first time, the data was not correct and there was an error with the index. This was because some of the lines were being run outside of the loop when it was supposed to be inside. Once this was corrected, there was a small error with coloring the positive returns green, so it took a few adjustments to correct the code so that it ran with no errors and gave the correct results.

In general, refactoring code is usually good practice because it makes your analysis more efficient and can save time (specifically with extremely large sets of data). It is always important for your code to run correctly and without error. So, this should be the first step when solving a problem or analyzing data. Then, once your code runs smoothly and you get some sort of result, you can go back and see if refactoring makes sense. It will not always make sense because there may be some cases where the data set is not too large or our original code runs in a small amount of time. In this case it may not be worth the time to refactor since the functionality of the code would not change and it already does what we want efficiently and effectively.

In this project, the most exciting advantages of the refactoring code are:
-faster run time which allows us to analyze more data or get results for current data quicker
-ability to analyze more than 12 stocks efficiently
	--Note: the refactored code could run thousands of stocks in a reasonable amount of time

The only disadvantage of the refactoring code is that it was slightly more tricky to write and it does not add any functionality to the original code written. I'd say there are more advantages than disadvantages because time and memory is expensive. Therefore, any code that takes less time and/or less memory to run is better to use.


### Skills used in analysis: VBA, excel
