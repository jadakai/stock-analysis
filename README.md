# Green Stock Analysis

### Project Overview

Steve has asked me to analyze a green stock for his parents to see if they should invest in it.  To do this, I used VBA in Excel to find the stocks total daily volume and the annual return. In the data provided there were additional stocks to compare with to see how the stock performed.  This comparison provided the best option on which stock to invest in.

### Purpose 

The purpose of this project was to find and use an efficent way of reviewing data from multiple stocks in VBA.  The initial code ran and produced the correct data.  However, there was a way to revise (refactor) the code to have it run in less time and become more efficient.  Although this project used 12 stocks refactoring code can be helpful with larger sets of data that can take more time to process.  

## Results

## Refactoring the Code

To make my code more efficient I needed to change the way code was nested within my loops by creating arrays.  Four arrays were created to achieve this; tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.  The tickers array was used to establish the ticker symbol of the stock and the variable called tickerIndex was used with the other three arrays. The variable TickerIndex was assigned to tickerVolumes, tickerStartingPrices, and tickerEndingPrices for each ticker symbol before running throught the set of data. This allows the data to be compiled faster than using the nested loop in the original code.

### Original Code

![](resources/VBA_Original_Code.r)

### Refactored Code

[]

## Run Time Videos for Original Code

### 2017
![](resources/original_2017.mov)

### 2018
[]

## Run Time Videos for Refactored Code

### 2017
[]

### 2018
[]

According to the results the refactored run times are faster and therefore more efficient.



