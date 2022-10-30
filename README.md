# Stock Market Analysis with VBA

## Overview of the Project

#### Purpose

Creating a workbook that will allow our client to analyze stocks with the click of a single button. Previously, we created code that accomplished this task for a limited data set. Currently, our client is interested in expanding his search to encapsulate the entire stock market for the past several years. To aid in this task, we will edit, or refactor, our existing code to improve its efficiency while handling a much larger data set.  

#### Goals

- To loop through all the data one time while collecting the same information that we previously accomplished.
- To produce script run time comparisons to demonstrate the new code's efficiency.

### Strategy

In the prior code, we were operating under one array, that of "tickers()". This was a String value to seperate out the individual stocks and combine their Volume, starting prices and ending prices through a series of repetitive loops.   

To catch all the data in one pass, we needed to create three additional arrays: tickerVolumes, tickerStartingPrices and tickerEndingPrices. These arrays will be numerical in value.

![arrays](https://github.com/aaron-ardell/stock-analysis/blob/main/Resources/Resources/1b.png)

To coordinate the data between these arrays, we'll need to create an Integer variable reflective of the ticker() array's String data. We've called this tickerIndex. 

![tickerIndex](https://github.com/aaron-ardell/stock-analysis/blob/main/Resources/Resources/1a.png)

Next, we'll use our single loop. First, we'll have the totalVolumes variable established so our loop picks up the cumulative volumes of all the stocks by name. Second, we'll utilitize If Then statements to allow our loop to determine the tickerStartingPrices and tickerEndingPrices. Notice the usage of the tickerIndex variable to determine the array's numeric variable.

![code](https://github.com/aaron-ardell/stock-analysis/blob/main/Resources/Resources/2a.png)

Finally, we're going to loop through our analysis worksheet to deliver our ticker names, cumulative stock volumes and the calculation of return values (tickerEndingPrice / tickerStartingPrices -1).

![4](https://github.com/aaron-ardell/stock-analysis/blob/main/Resources/Resources/refractored-4.png)

## Results

#### - Previous Results
![2017 Results](https://github.com/aaron-ardell/stock-analysis/blob/main/Resources/Resources/2017AllStockscells.png) ![2018 Results](https://github.com/aaron-ardell/stock-analysis/blob/main/Resources/Resources/2018allstockscells.png)

#### - Refactored Results
![2017 Results](https://github.com/aaron-ardell/stock-analysis/blob/main/Resources/Resources/2017refractoredcells.png) ![2018 Results](https://github.com/aaron-ardell/stock-analysis/blob/main/Resources/Resources/2018refractoredcells.png)

#### - Previous Elapsed Run Time for the Scripts

![2017 RTS](https://github.com/aaron-ardell/stock-analysis/blob/main/Resources/Resources/2017originalcodetimer.png) ![2018 RTS](https://github.com/aaron-ardell/stock-analysis/blob/main/Resources/Resources/2018originalcodetimer.png)

#### -  Refactored Elapsed Run Time For the Scripts

![2017 RRTS](https://github.com/aaron-ardell/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png.png) ![2018 RRTS](https://github.com/aaron-ardell/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png.png)

## Summary

The advantages of refactoring the code in this instance improved the speed in which the code could calculate and sort data for the client by limiting the number of times the code looped through every cell of data. Refactoring code can make it simpler to understand, more efficient in size and speed, more adaptive to the data set its applied to and simpler to understand. Writing the code from scratch would have added timely redundancy to the task and risked creating new errors to debug. The disadvantages of refactoring code would be the potential to introduce new error or bugs into existing code and the potentially time-consuming nature of the process. A risk vs. reward consideration should be carefully considered to deetermine if the potential outcome is worth the cost of resources. 

For this code there was a simple goal proposed by our client: streamline the code so it can be applied to the entire market worth of data. The original code was simple and effective at producing accurate results for the limited data set, but if the data set were increased the runtime created by the multiple loops through the cells would've hindered performance of the code. On the positive side, this meant that half the original code(timer, ticker array and cell formatting) were retained and only the central loop script would need to be refactored(variables, loops and If Then statements).   
