# stock-analysis


## Overview of Project
This project analyzes stock market data from twelve companies -- AY, CSIQ, DQ, ENPH, FSLR, HASI, JKS, RUN, SEDG, SPWR, TERP, and VSLR -- to find their Total Daily Volume and Yearly Returns for the years 2017 and 2018. Two sets of code were written to accomplish this, with the main difference between them being their use of for loops.

### Purpose
This purpose of this project is to see if there is any correlation between a stock's Total Daily Volumes and its Yearly Return, and whether or not a stocks Daily Volume is a good metric to use in predicting the profit of a stock. This project is also an exercise in comparing two different ways of coding for the same solution.

## Analysis

### Overview of Original Code
The first set of code 

### Overview of Refactored Code
The refactored code opted to 

## Results
These are the tables of the [2017 Stock Data](Resources/VBA_Challenge_2017.png) and the [2018 Stock Data](Resources/VBA_Challenge_2018.png), which show the Total Daily Volume and the Yearly Return for each of the 12 stocks.

It was hypothesized that a higher Daily Volume would correlate to a higher Yearly Return, however, looking at the two years side by side, there does not appear be any strong correlation between the Total Daily Volume and the yearly return. From the years 2017 to 2018, nearly half the stocks -- DQ, HASI, SEDG, TERP, and VLSR -- had an increase in Total Daily Volume, but a significant decrease in the Yearly Return. Even within the same year, Total Daily Volume does not reflect how high the Yearly Return might be.


## Summary

- Advantages and disadvantages of refactoring code in general

Generally, refactoring code is very useful for improving a code's efficiency and readability.

For the most part, disadvantages to refactoring code tends to be situational. Refactoring should not be done when there isn't enough time or funds to complete a project, because it can be quite time-consuming. (Stone, 2018) It can also be quite risky when the code for a program is very large, or when the refactorer isn't the same person who wrote the original code. New bugs could be introduced, which may harm the long term stability of the software or program. (Doug, 2008)

- Advantages and disadvantages of the original and refactored VBA script

The refactored code is significantly more efficient in calculating the Total Daily Volume and Starting and Ending Returns for each stock, because it only uses a single for loop instead of nesting one loop inside the other. So instead of looping through all of the data 12 times, once for each stock, it only has to loop through it once, which allows the calculations to run faster.

The disadvantage to both of the codes however, is that they require the data to be completely sorted by the type of stock. If any of the rows are scrambled so that the information for each stock is not grouped together, the calculations will fail and return incorrect information.

## Citations
Doug, T. (2008, September 28). Re: What are the limitations of refactoring? [Discussion post]. StackOverflow. https://stackoverflow.com/questions/146124/what-are-the-limitations-of-refactoring?rq=1

Stone, S. (2018, September 27). Code Refactoring Best Practices: When (and When Not) to Do It. altexsoft. https://www.altexsoft.com/blog/engineering/code-refactoring-best-practices-when-and-when-not-to-do-it/. 
