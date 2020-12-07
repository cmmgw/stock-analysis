# Green Stock Analysis

## Overview of Project
As fossil fuels get used up there will be more reliance on alternative energy production. Some sources of alternative energy include hydroelectric, wind, geothermal, solar and bioenergy and as such, there are many forms of green energy to invest in. This project analyzes thousands of green energy stocks from 2017 and 2018, through the utilization of Visual Basic for Applications (VBA), a programming language that interacts with excel by using macro code to automate analyses. 

### Purpose
The purpose of this project is to compare the stock performance between 2017 and 2018, along with the execution times of the original cript and the refactored script. By doing this, one will be able to determine whether refactoring code makes the VBA script run faster.


## Results
An in depth look at the performance of green stocks in 2017 and 2018, along with the detailed analysis outlined below, can be found **[here](https://github.com/cmmgw/stock-analysis/blob/main/VBA_Challenge.xlsm).** 

### Comparison of 2017 and 2018 Stock Performance
The tables below compare the performance of green stocks between 2017 and 2018. The “Ticker” column references the stock symbol assigned to a stock. The “Total Daily Volume” column highlights how many shares were traded per day. The “Return” column indicates the change in price of an investment over time and is represented by a percentage change. A positive return represents a profit and is shown on the table in green, while a negative return represents an economic loss and is shown on the table in red. Through the use of conditional formatting in VBA, the cell color in the “Return” column is automatically updated to green or red, depending on the outcome of the return, which simplifies interpretation of the data and enhances analysis.

![Resources/VBA_Challenge_Stock Performance_ 2017_and_2018](/Resources%2FVBA_Challenge_Stock_Performance_%202017_and_2018.JPG)

In taking a closer look at the analysis, the most visible difference is that in 2017 all but one stock (Ticker: TERP) had positive returns, while in 2018 all but two stocks (Tickers: ENHP and RUN) had negative returns. The data shows that two stocks (Tickers: ENPH and RUN) would have been good investments, as they had positive returns in 2017 and 2018. 

### Comparison of VBA Code Utilized to Run Analysis
To run the analysis, a detailed script of code had to be written in VBA. The original script was initially written and later refactored, with the intent to maximize the efficiency of the run time for the pop-up messages, after running analyses for 2017 and 2018.The original code included a nested for loop, which is essentially a loop inside of a loop, that tells the computer to repeat lines of code for as many loops outlined. In contrast, the refactored code was developed to loop through all the data one time in order to collect the same information as before. 

![Resources/VBA_Challenge_Original_vs._Refactored_VBA_Script](/Resources/VBA_Challenge_Original_vs._Refactored_VBA_Script.JPG)

### Comparison of Code Performance 
The original and refactored code both include a script that calculates how long the code takes to compile results and outputs the elapsed time in a message box. Based off of the results below, it is evident that the refactored code optimized efficiencies, by having shorter run times. 

#### Elapsed Run Time for 2017 Using Original vs. Refactored VBA Script

![Resources/VBA_Challenge_2017_Original](/Resources/VBA_Challenge_2017_Original.JPG)
![Resources/VBA_Challenge_2017_Refactored](/Resources/VBA_Challenge_2017_Refactored.JPG)

Difference in Run Time: 0.846436 seconds

#### Elapsed Run Time for 2018 Using Original vs. Refactored VBA Script 

![Resources/VBA_Challenge_2018_Original](/Resources/VBA_Challenge_2018_Original.JPG)
![Resources/VBA_Challenge_2018_Original](/Resources/VBA_Challenge_2018_Original.JPG)

Difference in Run Time: 0.8640139


## Summary

**- What are the advantages or disadvantages of refactoring code?** 
The main advantage of refactoring code is that it maximizes overall efficiency, through streamlining code and therefore reducing run times. A major disadvantage of having to rewrite code which already works is that it may not always be evident as to how to effectively do this, without running in to numerous run time errors which involve debugging the code.

**- How do these pros and cons apply to refactoring the original VBA script?**
The pros and cons outlined above directly applied to refactoring the original VBA script. Although the code had to constantly be debugged in order to effectively run, ultimately the code was more organized and ran more efficiently, allowing proper analysis of the data.   












