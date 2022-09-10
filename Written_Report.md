# stock-analysis

## Overview of Project: Explain the purpose of this analysis.
We are helping Steve find the total daily volume and yearly return for each stock.  

## Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
Daily volume for each stock is the total number of shares traded throughout the day (calculated with the sum of cells in column 8 in the specified year value for that stock (tickerVolume = tickerVolume + Cells(j, 8).Value) ) and the yearly return for each stock is the percentage difference in closing +-price of the stock from the beginning to the end of the year (calculated by checking the closing price in column 6 in the specified year value for that stock ((tickerEndingPrices = Cells(j, 6).Value/tickerStartingPrices = Cells(j, 6).Value) - 1) ).
The stocks did better in 2017 than they did in 2018 overall based on the yearly return percentages of each stock. The fefactored script ran ~.01 seconds slower than the original script.
![Original_Script_2018](https://user-images.githubusercontent.com/111719953/189505663-4066fc93-1aaf-4792-8c37-3bf0dd86f04c.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/111719953/189505666-77e05fc4-db4a-45ca-b60f-9656e7b6be04.png)


### Summary: In a summary statement, address the following questions.

- What are the advantages or disadvantages of refactoring code?
Refactoring code is much faster than starting from scratch but the placement of the code and the variable names could throw off the functionality and that needs to be adjusted when you implement it into your code.

- How do these pros and cons apply to refactoring the original VBA script?
With the refactored code, unlike the original, all of the formatting was part of the subroutine which shortened the script as a whole. I noticed the variable used to loop through the rows and grab the name-sets was different than where I needed it to be also the name of the variable of the inner-loop was not ideal so that also needed to be changed.
