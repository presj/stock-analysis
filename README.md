# stock-analysis
Module 2: VBA of Wall Street

Overview

This challenge involved the updating of previously analyzed stock data.  VBA codes were initially created to analyze the performance of an individual stock for a client, and then a list of stocks, by using a program flow that loops through a list of ticker symbols.  This code was additionally enhanced to run for any year, and then for even larger data sets.  Finally, a script was added to measure how long the code took to execute and output. By refactoring the last solution code, we sought to make the new code perform more efficiently.  Our aim through refactoring, was to use less steps and make the code easier for others to use.

Results

2017 and 2018 Stock Performance

Collectively, the 12 stocks analyzed outperformed in 2017 over 2018 with the average rates of return higher for 11 of the 12 stocks.  While returns were higher in 2017, the total daily volume was less than in 2018.  By contrast, TERP showed the only loss for 2017 while ENPH and RUN showed the only gains in 2018.  
 ![image](https://user-images.githubusercontent.com/100803302/158041133-a98dd268-832f-457a-86c7-02f41024ba0a.png)
 ![image](https://user-images.githubusercontent.com/100803302/158041141-0bdf0377-d5ab-4c9a-8e86-8630df48dd31.png)

A sampling of the code used to analyze this stock data is illustrated in these steps:
1a) Loop over all the rows in the spreadsheet
   For i = 2 To RowCount  
1b) Increase volume for current ticker         
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
1c) check if the current row is the first row with the selected tickerIndex      
If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then   
tickerstartingPrices(tickerIndex) = Cells(i, 6).Value     
End If
1d) check if the current row is the last row with the selected ticker
If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerendingPrices(tickerIndex) = Cells(i, 6).Value
End If
1e) Increase the tickerIndex.
If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerIndex = tickerIndex + 1
End If
Next i

Original and Refactored Script Execution Times

Before refactoring, the execution of code run times were 0.53125 seconds for the 2017 data and 0.51563 seconds for the 2018 data.  After refactoring, the execution run times are what follows:

![image](https://user-images.githubusercontent.com/100803302/158041169-60a8e7d3-78f4-419b-80d0-e0f983fbc09d.png)
![image](https://user-images.githubusercontent.com/100803302/158041179-56bc09d9-49f6-48b8-bbf9-4552735e5d00.png)

Summary

The refactoring of code presents some general advantages and disadvantages.  One advantage of refactoring code is that the script runs faster, allowing for the analysis of large datasets in a shorter time.  Another advantage is that restructuring code allows for ease of use without altering functionality.  Disadvantages of refactoring code includes introduction of new bugs and errors into the code, resulting in loss of functionality, and the time necessary to identify and correct problems. There were both pros and cons to refactoring the original VBA script.  First, the refactored code did run faster, but there was a loss in functionality.  Second, after taking more time to create a program that runs faster, I now have a program that continues to run after displaying the code run time.

