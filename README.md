

# Stock Analysis with VBA

## Overview of Project
Steve, who has a finance degree, wanted to help his to invest in DAQO New Energy Corporation stocks without knowing much about it. Before investing, Steve wanted to analyze other green energy stocks to help his parents to get better return by diversifying their investments. He made an excel data sheet with more than 3000 green energy stocks and he wanted us to help analyze the data with the help of VBA. 

### Purpose
Our purpose is to use an extension of Excel, VBA (Visual Basic Application) to automate tests and get better results for big size of data. VBA will help us to interact with excel; read and write the cells in the worksheet, make calculations, and perform analysis and reduce the chance of accidental errors. We will be using loops and conditions to see the yearly performances of 12 different company stocks. In this way not only we’ll be able to perform analysis on over 3000 data, but we’ll be giving reducing the chance of duplication and miscalculation. Our goal is to bring down the runtime of the code as much as possible to give quick result.


###RESULTS

By comparing only two years data for 12 stocks it is very veru difficult to decide where to invest as the results of the particular stock is totally different. Steve has to do lot more research on mor stocks information with several years data set.


!(png_Code before)[https://github.com/Ruma-T/stock_analysis/blob/b2f4c5f345c0eb06d7c0cfb372084add920aff8b/Code%20After.PNG]

!(png_Code After)[https://github.com/Ruma-T/stock_analysis/blob/b2f4c5f345c0eb06d7c0cfb372084add920aff8b/Code%20After.PNG]





### Analysis of All Stocks in 2017

•	Other than TERP, all stocks returns are very good, specially DQ where Steve's parents wanted to invest.
•	DQ's return value almost doubled in 2017.
•	Almost 33% performed over 100%.
•	Out of these 33%, 17% are close to double.
•	Only TERP's return is less than invested amount.





![png_All_Stocks_2017](https://github.com/Ruma-T/stock_analysis/blob/2915b05230cd963f3a4bf710113599f1d7fdc239/All_Stocks_2017.PNG)












### Analysis of All Stocks in 2018

•	ENPH and RUN performed well in 2018.
•	Rest of the stock returns are negative which means their return value is less than they invested.
•	Interestingly also performed very bad in 2018.
Compare All Stocks for 2017 & 2018







![png_All_Stocks_2018](https://github.com/Ruma-T/stock_analysis/blob/2915b05230cd963f3a4bf710113599f1d7fdc239/All-Stocks_2018.PNG)








Compare Runtime before and after Refactoring



Before



RunTime Before Refactoring

![png_2017 before refactoring]( https://github.com/Ruma-T/stock_analysis/blob/830b6388792fa4620ebf75309044892c0c204b26/2017%20before%20refactoring.PNG)





![png_2018_Before-Refactoring]( https://github.com/Ruma-T/stock_analysis/blob/830b6388792fa4620ebf75309044892c0c204b26/2018_Before-Refactoring.PNG)







Runtime After Refactoring



![png_VBA-Challenge_2017](https://github.com/Ruma-T/stock_analysis/blob/126c8877a4c5980c31232431adf7f0c1547c47dc/VBA_Challenge_2017.PNG)





![png_VBA-Challenge_2018](https://github.com/Ruma-T/stock_analysis/blob/126c8877a4c5980c31232431adf7f0c1547c47dc/VBA_Challenge_2018.PNG)







### Challenges and Difficulties Encountered

VBA was little difficult for me, specially learning this language with its pattern and specially learning to debug.
 Most of the times creating loops and putting the conditions together was difficult in the beginning.
The run time was much more than expected. I had figured out that nested loops take more times to run. So, I was looking for something that could help me lessen the runtime. I finally introduced counter in the index and made complex arguments to avoid nested loops. It drastically brought down the time.

Limitations of the Dataset
•	There are several stocks in green energy area, comparing all of them was not possible from the available data set.
•	There are different factors that affect the performance of the stocks, which could not be predicted early.
•	Comparing old data might not be the right choice to decide on a stock.
•	In data set all known factors used.
•	VBA can be utilized with lot of other type of data set.
•	Boolean is not used. 

- What are the possibilities to improve?
•	Compare different factors that affect the stock opening price and sells price. 
•	Period of investment is not considered in this data frame. 
•	How investing for shot time and longtime affects the performance could be included.
•	Volume of investment is also important to maximize the profit.

•	Summary
The advantages and disadvantages of refactoring code in general 
Advantage:
•	The code can be highly organized and clean.
•	We might get a better-quality code.
Disadvantage:
•	It is not recommended for big and stable program. 
•	Stable code might be harmed in the process of refactoring.
•	We might need to retest some functionality which might take lot of time.
•	To make it refactored, we might have to use overly complex arguments and continuing those arguments might be a problem if it is too big code.

The advantages and disadvantages of the original and refactored VBA script 
Advantage:
•	It took less time to run.
•	Accidental errors could be avoided.
•	We can take the help of creating a Buttons to run the work. 
•	Can work more efficiently.
•	Improved the logic.
•	Easier to understand for others.

Disadvantage:
•	It might not be very efficient to refactor for first time users
•	Organizing the code was a challenge as it needed to remove the nested loops
•	Time to refactor the code. 
•	Not able to add more functionality



