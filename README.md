-- VBA Stock-Analysis Data Analysis Written Report --
Written by: Sterling Miller



Overview of Project:

The purpose of this project is to consolidate and analyze stock market data to help our friend Steve curate financial recommendations of green (alternative energy) stocks for his parents to invest in. Our analysis combs 2017 and 2018 stock data, filters its contents by ticker, and analyzes its contents to calculate important data analytics metrics to assist Steve with his recommendations: such as stocks' Daily Trade Volume and Annual Return Percentage. In order to make these analyses easy to access and interpret, VBA (Visual Basic for Applications) Macros and Macro-Assigned Buttons have been implemented to run complex computations instantly with the simple click of a button, regardless of if the user accessing the workbook has their developer and/or VBA functionality enabled. With this analysis we strive to determine whether the "DQ" stock that Steve's parents originally wanted to invest in, is a lucrative/effective investment, and if not, what might be a suitable alternative.



Results:

The VBA code macro for the data analysis includes:
-A Timer for code performance measurement
-Data Table Formatting Script
-An Adaptive Worksheet Header based on the Year of Data being Analyzed
-Headers and Columns for three variables: "Ticker", "Total Daily Volume", "Return"
-An Array of green-stock tickers
-Iterative & Conditional loops for organizing, compiling, and assigning calculations to cells & columns


What we learned is that 2017's stock performance for the 12 green-stocks provided in the data was generally better than 2018's stock performance. Daily Trade Volume for green-stocks during 2017 was also considerably higher than the Daily Trade Volume of green-stocks during 2018
-11 out of 12 stocks had positive annual returns during 2017 

<img width="294" alt="Stock 2017" src="https://user-images.githubusercontent.com/87245870/147313424-a0f1eb68-c44f-4c1e-99b0-33074e061029.png">

-2 out of 12 stocks had positive annual returns during 2018.

<img width="294" alt="Stock 2018" src="https://user-images.githubusercontent.com/87245870/147313434-514998ec-4da2-4937-8d96-48c173e35c0d.png">

From this we can assume that green-stocks had a considerable dip or decline in growth & volume after 2017.

The "DQ" stock (the stock Steve's parents had been asking for analysis on), had seen a -62.6% return based on its performance in 2018. As such, alternative stock suggestions may be necessary.

The 2 stocks that remained positive in their returns in the 2018 data were "ENPH" and "RUN", returning 81.9% and 84% respectively. These could both be possible alternative stocks to recommend to Steve's parents based on the 2018 data analyses.


In reference to our code and our refactoring: 
- we've improved the analytics code's run performance from 0.5703 seconds for 2018's data to 0.1328 seconds

<img width="259" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/87245870/147313367-80a96b2d-698a-444e-8dcc-55cce8dc1db7.png">

- we've improved the analytics code's run performance from 0.5703 seconds for 2017's data to 0.1406 seconds

<img width="258" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/87245870/147313361-5f90ceb7-a530-49dd-83c2-447284ba676a.png">



Summary:

Q1) What are the advantages or disadvantages of refactoring code?
-By refactoring our code, we cleaned up its run time, making application of the code across large amounts of data faster and more applicable to further code amendments and machine-learning protocols. We've also re-edited our data to work beyond the scope of our single data set, making the coding structure/model applicable to any amount of stock data, thousands and beyond. Besides considerable time consumption, there doesn't seem to be any disadvantages, as the new refactored code runs faster, includes the formatting code in the same subroutine, and trims down most of the excess/extraneous code to make the end product more elegant and efficient.

Q2) How do these pros and cons apply to refactoring the original VBA script?
-The original VBA script separated the formatting, data compilation, and year inputting into separate functions. With the refactored version, everything exists within one subroutine, making the processes faster and more wholistic. Now instead of having 3 buttons (1 for Data Analysis, 1 for Formatting, and 1 for Clearing the Data), we only need 2 buttons: one of which asks the user what year he/she wants to analyze, analyzes the data, and formats the data to make it elegant and communicate more to the end user. Other than the fact that the refactoring took a considerably longer amount of time than not refactoring, I could see no noticeable cons, thus I have no comments on con applications. Refactoring has rich long term benefits!
