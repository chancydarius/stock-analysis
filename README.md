# stock-analysis

## Overview of Project: Explain the purpose of this analysis.
We are helping Steve find the total daily volume and yearly return for each stock. Daily volume for each stock is the total number of shares traded throughout the day (calculated with the sum of cells in column 8 in the specified year value for that stock (tickerVolume = tickerVolume + Cells(j, 8).Value) ) and the yearly return for each stock is the percentage difference in price of the stock from the beginning to the end of the year (calculated by checking the closing price in column 6 in the specified year value for that stock ((tickerEndingPrices = Cells(j, 6).Value/tickerStartingPrices = Cells(j, 6).Value) - 1) ). 

## Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
The stocks did better in 2017 than they did in 2018 overall based on the yearly return percentages of each stock. The fefactored script ran ~.02 seconds slower than the original script.

### Summary: In a summary statement, address the following questions.
- What are the advantages or disadvantages of refactoring code?
Refactoring code is much faster than starting from scratch but the placement of the code and the variable names could throw off the functionality and that needs to be adjusted when you implement it into your code.

- How do these pros and cons apply to refactoring the original VBA script?
I noticed what variable used to loop through the rows and grab the name-sets was different than where I needed it to be also the name of the variable of the inner-loop was not ideal so that also needed to be changed.
