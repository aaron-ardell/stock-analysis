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

To coordinate the data between these arrays, we'll need to create an Integer variable reflective of the ticker() array's String data. We've called this tickerIndex. 

Next, we'll use our single loop. First, we'll have the totalVolumes variable established so our loop picks up the cumulative volumes of all the stocks by name. Second, we'll utilitize If Then statements to allow our loop to determine the tickerStartingPrices and tickerEndingPrices. Notice the usage of the tickerIndex variable to determine the array's numeric variable.

Finally, we're going to loop through our analysis worksheet to deliver our ticker names, cumulative stock volumes and the calculation of return values (tickerEndingPrice / tickerStartingPrices -1).

## Results

#### - Previous Results


#### - Refactored Results


#### - Previous Elapsed Run Time for the Scripts



#### -  Refactored Elapsed Run Time For the Scripts


### Summary

The advantages of refactoring the code in this instance improved the speed in which the code could calculate and sort data for the client by limiting the number of times the code looped through every cell of data. Refactoring code can make it simpler to understand, more efficient in size and speed, more adaptive to the data set its applied to and simpler to understand. Writing the code from scratch would have added timely redundancy to the task and risked creating new errors to debug. 


What are the advantages or disadvantages of refactoring code?

How do these pros and cons apply to refactoring the original VBA script?

What are the advantages or disadvantages of refactoring code?




For your written analysis, be sure to use complete and coherent sentences. Your written analysis should contain three sections, which cover the following:

Overview of Project: Explain the purpose of this analysis.
Results: Using images and examples of your code, compare the stock performance between 2018 and 2018, as well as the execution times of the original script and the refactored script.
Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
Deliverable 2 Requirements
Structure, Organization, and Formatting Requirements (8 points)
The written analysis contains the following structure, organization, and formatting:

There is a title, and there are multiple paragraphs (2 pt).
Each paragraph has a heading (2 pt).
There are subheadings to break up text (2 pt).
Links are working, and images are formatted and displayed where appropriate (2 pt).
Analysis Requirements (12 points)
The written analysis has the following:

Overview of Project
The purpose and background are well defined (2 pt).
Results
The analysis is well described with screenshots and code (4 pt).
Summary
There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).
