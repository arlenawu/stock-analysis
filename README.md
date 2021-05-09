# stock-analysis


## Overview of Project
This project analyzes stock market data from twelve companies -- AY, CSIQ, DQ, ENPH, FSLR, HASI, JKS, RUN, SEDG, SPWR, TERP, and VSLR -- to find their Total Daily Volume and Yearly Returns for the years 2017 and 2018. Two sets of code were written to accomplish this, with the main difference between them being their use of for loops.

### Purpose
This purpose of this project is to see if there is any correlation between a stock's Total Daily Volumes and its Yearly Return, and whether or not a stocks Daily Volume is a good metric to use in predicting the profit of a stock. This project is also an exercise in comparing two different ways of coding for the same solution -- a base set of code and its improved refactored version.

## Analysis
### Organization of the Raw Data
The raw data set is divided into two separate datasets, one for each year, 2017 and 2018. Both datasets are sorted first by stock name, and then by date from oldest to newest.

### Overview of Original Code
The code overall is fairly straightforward. First, a table was created and formatted on a new sheet which is designed to hold all the final calculated information of the Totaly Daily Volumes and the Yearly Returns for each stock. Then, an array of strings was created to act as a list of the names of each stock. Then two for loops were created -- one to iterate through the array of stock names, and the other to iterate through the dataset for the specified year.

So, for each stock name, the code would iterate through the dataset in search of that stock name and all of its data. It wuold calculate the Total Daily Volume by summing the daily volume from every single data point of that stock, and it would calculate the Yearly Return by dividing the stock's Ending Price (the price of the last data point listed for that stock) by the Starting Price (the price of the first data point listed for that stock) and subtracting 1 from the result. And then, the Total Daily Volume and Yearly Return would be printed into the table on the new sheet.

Below is a snippet of the for loops used in the code.

    Note: The term 'ticker' means a stock's name
    ---------------------------------------------
    'loop through each of the 12 tickers
    For i = 0 To 11
    
        Worksheets(yearValue).Activate
        ticker = tickers(i)
        totalVolume = 0
        
        'Loop through each row in specified sheet
        For j = 2 To RowCount
        
            'Find total volume of ticker
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
            
            'Find starting price for ticker
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            
            'Find ending price for ticker
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingPrice = Cells(j, 6).Value
            End If
            
        Next j
        
        'Output data for current ticker
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1
        
    Next i

To finish things off, a button was placed on the spreadsheet that the user could click on to begin the analysis, and upon cicking the button a dialogue box was set to pop up to ask the user which year they would like to run the analysis on (2017 or 2018).

### Overview of Refactored Code
The refactored code focused on reducing the number of iterations through the dataset that the for loops in the original code made. Instead of iterating through the data once for each of the twelve stock names, or 'tickers', for a total of twelve times, the refactored code only needs to interate through all the data *once* while calculating the Total Daily Volume and the Yearly Returns for all twelve stocks at the same time.

The simplification process of the for loops was first begun by creating three new arrays -- tickerVolumes, tickerStartingPrices, and tickerEndingPrices. Each of these three arrays contain twelve entries to represent each of the 12 stocks, or 'tickers', and to hold the values of their Total Daily Volumes, Starting Prices, and Ending Prices.

    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

Next, the previous set of nested for loops from the original code was eliminated and replaced with a single for loop that would iterate once through the dataset. Within this for loop, the three arrays, tickerVolumes, tickerStartingPrices, and tickerEndingPrices, were set to have their values modified as the code iterated its way through the dataset. By the end of the loop, all three arrays should contain the correct Total Daily Volumes, Starting Prices, and Ending Prices for each of the twelve stocks/tickers.

    'Loop over all the rows in the spreadsheet
    For i = 2 To RowCount
        
        'Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        'Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        'Check if the current row is the last row with the selected ticker
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            'Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i

Finally, a for loop was written to iterate through the three arrays simultaneously, and print out the stock names, Total Daily Volumes, and Yearly Returns into a neat table in the spreadsheet.

    'Loop through arrays to output the Ticker (Stock Name), Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(i + 4, 1).Value = tickers(i)
        Cells(i + 4, 2).Value = tickerVolumes(i)
        Cells(i + 4, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1

    Next i

## Results
These are the tables of the [2017 Stock Data](Resources/VBA_Challenge_2017.png) and the [2018 Stock Data](Resources/VBA_Challenge_2018.png), which show the Total Daily Volume and the Yearly Return for each of the 12 stocks.

It was hypothesized that a higher Daily Volume would correlate to a higher Yearly Return, however, looking at the two years side by side, there does not appear be any strong correlation between the Total Daily Volume and the yearly return. From the years 2017 to 2018, nearly half the stocks -- DQ, HASI, SEDG, TERP, and VLSR -- had an increase in Total Daily Volume, but a significant decrease in the Yearly Return. Even within the same year, Total Daily Volume does not reflect how high the Yearly Return might be.


## Summary

- Advantages and disadvantages of refactoring code in general

Generally, refactoring code is very useful for improving a code's efficiency and readability, which can improve performance and clean up bad structures in the code such as redundant or unused code. (Cuelogic, 2014) It's a way of keeping the code as simple and clean as possible for long term maintenance, and is important to do before adding any major new features or changes so that the addition of the new code does not risk making things too convoluted and messy. (Stone, 2018) It's also helpful in debugging and preventing further defects and bugs from being created. (Ershad, 2017)

For the most part, the disadvantages to refactoring code tends to be situational. Refactoring should not be done when there isn't enough time or funds to complete a project, because it can be quite time-consuming. (Stone, 2018) It can also be quite risky when the code for a program is very large, or when the refactorer isn't the same person who wrote the original code. New bugs could be introduced, which may harm the long term stability of the software or program. (Doug, 2008) It is important to plan carefully about when to refactor code and the extent to which it should be done before beginning. (Stone, 2018)

- Advantages and disadvantages of the original and refactored VBA script

The refactored code is significantly more efficient in calculating the Total Daily Volume and Starting and Ending Returns for each stock, because it only uses a single for loop instead of nesting one loop inside the other. So instead of looping through all of the data 12 times, once for each stock, it only has to loop through it once, which allows the code to run faster.

The disadvantage to both codes however, is that they require the data to be completely sorted by the type of stock and then by their date from oldest to newest. If any of the rows are scrambled so that this information is no longer in order, the calculations will fail and return incorrect information.

## Citations
Cuelogic. (2014, August 27). *What is refactoring and Why is it important?* Cuelogic Technologies Pvt. Ltd. https://www.cuelogic.com/blog/what-is-refactoring-and-why-is-it-important. 

Doug, T. (2008, September 28). *Re: What are the limitations of refactoring?* [Discussion post]. StackOverflow. https://stackoverflow.com/a/146143.

Ershad, G. M. (2017, January 9). *Pros And Cons Of Code Refactoring.* C# Corner. https://www.c-sharpcorner.com/article/pros-and-cons-of-code-refactoring/. 

Stone, S. (2018, September 27). *Code Refactoring Best Practices: When (and When Not) to Do It.* altexsoft. https://www.altexsoft.com/blog/engineering/code-refactoring-best-practices-when-and-when-not-to-do-it/. 
