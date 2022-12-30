# Stock-analysis

## Overview.
The purpose of this project is to refactor the code of Module 2 that collect information about stock market during years of 2017 – 2018; and determinate if this process will make the VBA script run faster.

## Background
The module 2 code calculates the total daily volume and yearly return for stock. Daily volume is the total number of shares traded throughout the day; it measures how actively a stock is traded. The yearly return is the percentage difference in price from the beginning of the year to the end of the year. 
The results are showed in a table, it is formatting to determine stock performance immediately. We are looking to loop through all the data one time to collect the same information but efficiently.

## Results
After imported VBA_Challenge.vbs to script to the Microsoft Visual Basic editor, I reused the code of set variables the input box, chart headers, ticker array, and to activation of the worksheet. Then , I follow the steps to create a tickerIndex , set arrays, create a loop to thought all data and conditionals for the starting and ending prices.

![image](https://user-images.githubusercontent.com/120151872/210043332-73b6c4c1-c17f-4fca-9e7a-c81733155267.png)

As a result in 2017 stock market had a better performance. 92% of stocks give a positive return compare with 17% of 2018.  Total daily volume in 2018 exceed  4% the previous year however  only 2 stocks give return that represent 33% of stocks traded. 

![2017](https://user-images.githubusercontent.com/120151872/210043674-becd8d07-d5d8-4e0b-9649-fe3345002c3b.PNG)     

![tabla 2008](https://user-images.githubusercontent.com/120151872/210043716-a1028d74-3744-483f-a4ae-dcd4766172f6.PNG)


Regarding the execution time, the refactored script runs in less time than the original .
