# Stock-analysis

## Overview of Project
To examine and refactor VBA codes for this stock analysis by looping through all the data one time in order to collect the same information as before, and determine whether refactoring has successfully made the VBA script run faster.Finally, present a written analysis that explains findings and show results in three arrays of tickers, total volumes and returns. Steps are as follows:

### Steps:

A tickerIndex was created to equal to zero, which subsequently was looped through the rows to access the tickers array and the three output arrays created later, namely, tickerVolumes, tickerStartingPrices, and tickerEndingPrices. 

Then a nested loop was used to increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker.

Used an if-then statement to check if the current row is the first/last row with the selected tickerIndex. then assign the current starting/closing price to the tickerEndingPrices variable.

Wrote script to increase the tickerIndex if the next row’s ticker doesn’t match the previous row’s ticker. Finally. used a for loop to loop through the arrays (tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to output the “Ticker,” “Total Daily Volume,” and “Return” columns in the spreadsheet.

## Results
As displayed in the two screenshots, refactoring did make the process run more efficiently with less time used

