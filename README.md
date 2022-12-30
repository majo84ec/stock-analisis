# Stock-analysis.

## Overview.
The purpose of this project is to refactor the code of Module 2 that collect information about stock market during years of 2017 – 2018; and determinate if this process will make the VBA script run faster.

## Background.
The module 2 code calculates the total daily volume and yearly return for stock. Daily volume is the total number of shares traded throughout the day; it measures how actively a stock is traded. The yearly return is the percentage difference in price from the beginning of the year to the end of the year. 
The results are showed in a table, it is formatting to determine stock performance immediately. We are looking to loop through all the data one time to collect the same information but efficiently.

## Results.
After imported VBA_Challenge.vbs to script to the Microsoft Visual Basic editor, I reused the code of set variables the input box, chart headers, ticker array, and to activation of the worksheet. Then , I follow the steps to create a tickerIndex to access all stock ticker names , set output arrays, create a loop to  go through all data and the conditionals for the starting and ending prices.

![image](https://user-images.githubusercontent.com/120151872/210043332-73b6c4c1-c17f-4fca-9e7a-c81733155267.png)

As a result in 2017 stock market had a better performance. 92% of stocks give a positive return compare with 17% of 2018 where ENPH and RUN stocks had the best performance.  Total daily volume in 2018 exceed 4% the previous year and it represent 33% of stocks traded even though  only to 2 years were analyzed, the ENPH and RUN stocks could tend to be worthy investment in the future. 

![2017](https://user-images.githubusercontent.com/120151872/210043674-becd8d07-d5d8-4e0b-9649-fe3345002c3b.PNG)     

![tabla 2008](https://user-images.githubusercontent.com/120151872/210043716-a1028d74-3744-483f-a4ae-dcd4766172f6.PNG)


Regarding the execution time, the refactoring code runs the macro in less time than the original see the image below.

![COMPARACION](https://user-images.githubusercontent.com/120151872/210092242-7447ea81-0c09-46d1-b7ac-3ee23d4e90d7.png)

## Summary.

### Advantages and disadvantages of refactoring code.

Refactoring makes a clean code,  easy to understand, extend and maintain in the future by eliminate unnecessary complexity or features, in this process is possible to find some bugs that help improve the code quality. However , there are risks to edit a code when the application is big and the developers do not understand what's all about and it this process become expensive and time consuming so it is necessary to evaluate if it process will be worthy.

### Advantages and disadvantages of the original and refactored VBA script.
The disadvantage of original code is used a  'ticker(i)' variable in a for loop to read the data of each index position and it increased the execution time of AllStockAnalysis macro. The refactoring VBA script eliminate this excessive iterations and it will be helpful to other users understand what is behind of macro because it is more concise and organized . 
The first version  required the ticker array list in the same order as the ticker appears in the dataset in order to aviod error in the analysis. 
 , no matter size project  if refactoring is not doing properly could introduce new bugs and errors.

