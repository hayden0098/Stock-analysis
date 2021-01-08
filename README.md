# Stock-analysis

## Overview of Project
The Stock data sets from Wall Street that are using to help Steve analysis Alternative Green Energy stock and a specific stock that called Daqo.
### Purpose
Write clear and readable VBA code to calculate the total daily volume that the total number of shares traded throughout the day; yearly return for stocks that the percentage difference in price from the beginning to the end of the year. The volume measures how actively stock is traded and the return shows the percentage of increase or decrease that indicate how much investment grew or shrunk by the end of year.

After some work, we see from the above table that Daqo(DQ) stock is negative return in last year, which mean the return is shrunk at the end of year and this is the case that would not consider as the one worth to invest. So we will anaylze mutiple stocks in whole list to find some other better choice to invest in.

## Results
In 2017, Many stocks were have the positive return especially DQ(with volume 35,796,200), SEDG(with volume 206,885,200) and ENPH(with volume 221,772,100); TERP(with volume 139,402,800) is the only stock have negative return.

In 2018, ENPH(with volume 607,473,500) and RUN(with volume 502,757,100) were have the positive return; The highest negative return stocks were DQ(with volume 107,873,900), JKS(with volume 158,309,000) and SQWR(with volume 538,024,300).
Lets combine 2017, 2018 all stocks data sets, we can see ENPH(with volume +385,701,400 from last year) and RUN(with volume +235,075,800 from last year) are stocks that keep growthing in two years. These 2 stocks could be the option to consider inviest in.
while working on VBA, we have refactored the All Stocks Analysis subrountine code in order to make the VBA script run faster. We figured out a way to use one loop and tickerIndex variable as index instead of nested loop, this can make computer use les memory by improving the logic of the code.



## Summary

1. ###### What are the advantages or disadvantages of refactoring code?

2. ###### How do these pros and cons apply to refactoring the original VBA script?
