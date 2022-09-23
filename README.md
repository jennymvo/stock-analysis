# stocks-analysis

## Overview of Project
The project consists of creating macros that compiles the percentage of returns of different stocks from the years 2017 and 2018. 

### Purpose
The purpose of this project is to create macros that will:
    1. loop through rows of data
    2. auto-populate information and calculations
    3. automatically format worksheets
    4. learn how to refactor code
    
## Results and Analysis 

### Overview of the VBA script

The script first intitializes a timer before running the script, and when it finishes, the timer is ended and a message box appears for the calculated total time that it took to complete the analysis. 

![start timer](https://github.com/jennymvo/stock-analysis/blob/main/images/startTime.png)

![end timer](https://github.com/jennymvo/stock-analysis/blob/main/images/endTime.png)

To analyze the returns from the years 2017 and 2018, the script takes in an excel workbook that has columns of data that contains:
    1. Tickers
    2. Dates 
    3. Open 
    4. High
    5. Low
    6. Close
    7. Adjusted Close
    8. Volume

![Example from the 2018 Worksheet](https://github.com/jennymvo/stock-analysis/blob/main/images/2018_worksheet.png)

From the choice of 2017 or 2018, the script activates the output sheet that populates the data. It first creates headers of Tickers, Total Daily Volume, and return, then it initialized the array of tickers. 

![Headers and tickers](https://github.com/jennymvo/stock-analysis/blob/main/images/init_tickers.png)

It activates the data worksheet for the respective year, and output arrays are created to store the values. 

![Worksheet activation and array initialization](https://github.com/jennymvo/stock-analysis/blob/main/images/init_worsksheet_arrays.png)

It sets the tickerIndex to 0 in order to loop through the different tickers. It loops through each row and checks for the ticker, and finds the value of the volume and adds it to the tickerVolume array. Then it makes sure that the current row is the first row with the selected tickerIndex number that corresponds to the ticker, and if it is the case then it adds the starting price value to the tickerStartingPrice array, and does the same for the tickerEndingPrice array. It then increases the tickerIndex to do the same loop for the next ticker until it finishes all the rows. 

![Looping through the worksheet](https://github.com/jennymvo/stock-analysis/blob/main/images/loop.png)

When it finishes, it loops through all the arrays to output the ticker, total daily volume, and calculates the returns by dividing the starting price by the ending price and then minus 1. 

![Output Tickers, Total Daily Volumes, and Returns](https://github.com/jennymvo/stock-analysis/blob/main/images/output.png)

The next set of lines formats the sheet for asthetics and conditionals for negative and positive returns showing red and green respectively. 

### Results
The results are shown below:

![2017](https://github.com/jennymvo/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

![2018](https://github.com/jennymvo/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

## Summary

### Advantages and Disadvantages of refactoring code
An advantage of refactoring the code allows for the code to be easier to understand, more simple, and easier to modify. 

A disadvantage of refactoring code is that it can be time consuming, especially if the code needs to be re-written from scratch, is difficult to understand, or contains many dependencies that are easy to break. If the refactoring does not work completely well, there is a chance that new bugs are introduced into the code. 

### Advantages and Disadvantages of the original and refactored VBA script

The original starter code was able to ask for the year in which the user would like to analyze the total daily volumes and returns from all the differnt tickers. It was able to generate the headers, and a conditional of less than a positive integer in the returns column which was just empty wells. 

![Before Refactoring the code](https://github.com/jennymvo/stock-analysis/blob/main/images/beforeRefectoring.png)

After refactoring the code, it was able to generate data for each ticker and fill in the columns as shown in the results section, which is a huge advantage from it generating nothing. 

A disadvantage of the code, is it can only calculate information from the worksheets that exist in the workbook, which is limited in itself. There are only two different years to choose from, and there is information for only twelve tickers. It would be useful if the code is able to compile information for any existing tickers that the user would like for any year from an easily accessable source. 
