# stock-analysis

## Overview of Project: 
### *Explain the purpose of this analysis.*
#### The purpose of this Module and Project is to learn how to code in VBA (Visual Basic for Applications) Macros in Excel using Microsoft Office applications, and thus learning some of the building blocks of programming. 
####
#### VBA code was written to create VBA Macros to help Steve with the analysis of a set of green energy stocks contained within the Excel dataset (green_stocks.xlsx) he has given me. VBA was used to automate the tasks he requested to be used with any stock and with multiple years returning the Stock Ticker Name, Stock Daily Volume, and Average Daily Return. 

##
 ## Results: 
### *Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.*
#### 
#### For the DQ Stock that Steve’s Parents are interested in trading in, it had the highest return at 199.4% in 2017 ![alt text](https://github.com/Chipsheffield/stock-analysis/blob/main/Resources%20for%20Module%202%20Challange/VBA_Challenge_2017.png) and the worst Return at -62.6% in 2018 (See Image of 2018 Chart Analysis at ![alt text](https://github.com/Chipsheffield/stock-analysis/blob/main/Resources%20for%20Module%202%20Challange/VBA_Challenge_2018.png ) Based on DQ’s history I would not recommend it for purchase. The two stocks I would recommend instead are ENPH which went from 129.5% in 2017 to 81.9% in 2018, or RUN which grew from 5.5% to 84.4%. 
####
#### In terms of performance the 2017 script decreased from Original at 0.6796875 seconds ( See Image of 2017 green_stock Chart Analysis at https://github.com/Chipsheffield/stock-analysis/blob/main/Resources%20for%20Module%202%20Challange/green_stocks%202017.png   ) to 0.1210938 seconds for the Refactored, and the 2018 script decreased from Original ( See Image of 2018 green_stock Chart Analysis at https://github.com/Chipsheffield/stock-analysis/blob/main/Resources%20for%20Module%202%20Challange/green_stocks%202018.png   ) at 0.671875 seconds to 0.1054688 seconds for the Refactored, thus increasing the efficiency of the code. 

##
## Summary: 
### In a summary statement, address the following questions.
####
### *What are the advantages or disadvantages of refactoring code?*
#### The advantage of refactoring code is keeps you from having to reinvent the wheel when writing code for projects that have similar specs, requirements, or outcomes. Refactoring allows you to cut and paste previously written code and adapt it to the project at hand decreasing coding time. A disadvantage is if you are not careful in refactoring your code it could create a unusable product that will not run or return the results or data you are looking for. Another advantage of using the Refractoring code was as seen in the Results section above, in that the run time decreased for both 2017 and 2018.  
####
### *How do these pros and cons apply to refactoring the original VBA script?*
#### I was able to save a lot of time copying over the AllStocksAnalysisRefactored subroutine file to the Macro and since I didn’t have to code all the formatting, inputs, and reporting. What gave me problems was trying to figure out how to loop with only one variable when Module 2 gave an example with two. After I was able to code out the whole entire Macro it kept crashing until I figured out, I had an error in the code that increases the tickerIndex by using Cells(i-1,1) instead of Cells(i+1,1). ![alt text](https://github.com/Chipsheffield/stock-analysis/blob/main/Resources%20for%20Module%202%20Challange/VBA_Challenge%20%20Code%20Screenshot%203.png)       (See screenshot at https://github.com/Chipsheffield/stock-analysis/blob/main/Resources%20for%20Module%202%20Challange/VBA_Challenge%20%20Code%20Screenshot%203.png)

