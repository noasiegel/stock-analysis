# stock-analysis

## Overview of Project

The purpose of this analysis was to understand how to edit, or refactor, a past solution we uncovered during the coursework. The goal was the loop through all the data one time in order to collect the same information that we did while working through the module lessons. Ultimately, we wanted to understand if by refactoring our code we could make the VBA script run faster. Fundamentally, the purpose of the analysis was to give Steve a tool to easily and efficiently understand the stock market prices in order to create a robust, diversified portfolio for his parents.

## Results
### Stock Performance

As shown in the images below, 2017 was a much more successful year for the 12 stocks we have analyzed. TERP is the only stock that had a negative return (-7.2%). In 2018, only two stocks returned a positive value, ENPH and RUN, at 81.9% and 84%, respectively. Although 2018 was a net loss in terms of investments, the RUN stock outperformed itself from 2017 to 2018. In 2017, this stock only had a 5.5% return, where in 2018, this stock had an 84% return rate. This would be a wise stock to invest in. Similarly, although the ENPH stock decreased its return rate from 10`7 to 2018 by 47.6% (129.5% in 2017 to 81.9% in 2018), this would also be a wise stock to invest in because it is one of only two stocks that had a positive return over these two years. 

<img width="230" alt="All_Stocks_2017" src="https://user-images.githubusercontent.com/110838228/188037526-b246da07-377e-42b0-bad6-40ec8c081343.png">
<img width="230" alt="All_Stocks_2017" src="https://user-images.githubusercontent.com/110838228/188037529-ca606a8e-4153-48fe-9b3d-12f4b708668c.png">

Although the same code is used to analyze the all years of stock data, the formatting section of our refactored code aides us tremendously in analyzing stock performance as compared to our original AllStocksAnalysis macro. By adding visual elements to the table, namely green and red cells, all people of varying levels of analytical skills are able to discern success from failure. The following code is what we used to create this effect:

    Worksheets("All Stocks Analysis").Activate

    Range("A3:C3").Font.FontStyle = "Bold"
    
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    Range("B4:B15").NumberFormat = "#,##0"
    
    Range("C4:C15").NumberFormat = "0.0%"
    
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For I = dataRowStart To dataRowEnd
        
        If Cells(I, 3) > 0 Then
            
            Cells(I, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(I, 3).Interior.Color = vbRed
            
        End If
        
    Next I
    
All formatting before the For loop is setting up the headers of the table with more aesthetically pleasing styles. While it looks better, it also helps the viewer understand that "Ticker", "Total Daily Volume" and "Return" are the elements we are sharing a story about. All code within the For loop is assigning each cell to be filled with either Green or Red, according to if the value is greater than or less than 0. 

### Execution Times

Below are screenshots of the execution time of our original script, AllStocksAnalysis. This script ran in 0.281 and 0.304 seconds, for 2017 and 2018, respectively.

<img width="274" alt="VBA_Challenge_2017_BeforeRefactor" src="https://user-images.githubusercontent.com/110838228/188039702-c961e1ac-24ef-4d6f-91e5-2623595092cf.png">
<img width="274" alt="VBA_Challenge_2017_BeforeRefactor" src="https://user-images.githubusercontent.com/110838228/188039684-03dde608-8c62-4a52-98be-ee88a204cd5a.png">

After refactoring, or editing this original code, we were able to make it more efficient. As you can see below, the AllStocksAnalysisRefactored is 0.2109 seconds faster for the 2017 stock data, and 0.2343 seconds faster for the 2018 stock data.

<img width="278" alt="VBA_Challenge_2017_AfterRefactor" src="https://user-images.githubusercontent.com/110838228/188040110-a77afec2-177a-41c0-b275-0e0309c51290.png">
<img width="268" alt="VBA_Challenege_2018_AfterRefactor" src="https://user-images.githubusercontent.com/110838228/188040125-5b633657-6217-4f50-8742-3a92489dea2d.png">

Our refactored code is more efficient simply because of how we ordered, or logically structured, the code. In our original script, we were asking the computer to get the outputs for each column within multiple For loops. We asked the computer to loop through all the tickers and set their volume to 0. Once that step was achieved, we asked the data to loop through all the rows - within this loop we asked the computer to get the total volume for the current ticker. The next loop acted similarly, but instead got the starting price for the curent ticker. The next loop got the ending price for the current ticker. Finally, we output all the findings from the loops to populate in the ceels, including Total Daily Volume and Return. The code looks like this:

    '4) Loop through tickers
   For I = 0 To 11
       Ticker = tickers(I)
       totalVolume = 0

       '5) loop through rows in the data

       Worksheets("2018").Activate
       For J = 2 To RowCount

           '5a) Get total volume for current ticker

           If Cells(J, 1).Value = Ticker Then

               totalVolume = totalVolume + Cells(J, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(J - 1, 1).Value <> Ticker And Cells(J, 1).Value = Ticker Then

               startingPrice = Cells(J, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(J + 1, 1).Value <> Ticker And Cells(J, 1).Value = Ticker Then

               endingPrice = Cells(J, 6).Value

           End If
       Next J
       '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + I, 1).Value = Ticker
       Cells(4 + I, 2).Value = totalVolume
       Cells(4 + I, 3).Value = endingPrice / startingPrice - 1

   Next I


In our refactored code, we asked the computer to do the same exact thing, but we used a tickerIndex variable to access the correct index across our four arrays. Essentially, by using tickerIndex, we are telling the code exactly where to look to find what we need. Without using tickerIndex, the program is searching through every row and column to find the desired output. The code looks like this:

1a) Create a ticker Index
    
    Dim tickerIndex As Long
    tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickerVolume(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
       For I = 0 To 11
       tickerVolume(I) = 0
       
    Next I
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    For I = 2 To RowCount
    
        '3a) Increase volume for current ticker
    
        tickerVolume(tickerIndex) = tickerVolume(tickerIndex) + Cells(I, 8).Value
     
        
        '3b) Check if the current row is the first row with the selected tickerIndex(stock).
        
        
        
        If Cells(I - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(I, 6).Value
        End If
        
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        
         If Cells(I + 1, 1).Value <> tickers(tickerIndex) Then
         tickerEndingPrices(tickerIndex) = Cells(I, 6).Value
         
         
            '3d Increase the tickerIndex.
            
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next I
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For I = 0 To 11
    
        
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + I, 1).Value = tickers(I)
        Cells(4 + I, 2).Value = tickerVolume(I)
        Cells(4 + I, 3).Value = tickerEndingPrices(I) / tickerStartingPrices(I) - 1
        
        
    Next I


## Summary

The advantages of refactoring code are creating a more efficient or easily understood script. Refactoring can mitigate redundanices in code excessive/confusing code. The disadvantages of refactoring code are accidentally introducing bugs or areas that are more susceptible to breaking. Additionally, if you're working with a team to refactor code, coordination efforts could increase.

These pros directly apply to refactoring the original VBA script because it is evident that the script is more efficient based on our time output box. The cons directly apply to refactoring the original VBA script because while writing the code, it was more common (for me) to run into debugging errors than on the prior macro.

