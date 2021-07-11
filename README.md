# Stock Analysis

## Project Overview
Our client Steve had a large data sheet of the performance of 12 different stocks, which contained information like the stock ticker, opening and closing price over a twelve month periods, and to total trade volume of each stock. We were given this information for both 2017 and 2018. Our task was to analyze this data, which we did using Visual Basic for Applications (VBA) and excel Macro's.
### Purpose
Using VBA, we attempted to use VBA to write code to complete the analysis for us. This allows us to reuse the functioning code for any similar data sets in the future that our client would like to analyze. This also ensures that human error will not come into play with the analysis, as the code will make sure the analysis is run correctly for each stock we wish to learn more about. We wrote VBA code to run through every line of the data sheet to total the trade volume of each specific stock, as well as calculate the percentage return from each stock.
### Background
We have been asked to analyze a variety of Green Energy stocks for a client who is looking to advise his parents on which companies and stocks have provided the best return in the green or renewable energy space. They have chosen a company they want to invest in, but our client wants more information about the stock they have chosen as well as other companies in a similar space. We used VBA scripts to automate this analysis. 
## Results
To complete the analysis, we wrote a few different subroutines. First we analyzed just one stock, ticker:DQ, as almost proof of concept to show that our code worked well for the one stock. We first used the Dim function to assign variables for Total Daily Volume (totalVolume), stock starting price (startingPrice), and stock ending price (endingPrice). Then we set the totalVolume = 0, and created a For loop to analyze all DQ trades. The first loop we ran was from row 2 to row end, which was the entire data set minus the headers. Inside this For loop we had three conditionals, listed here.

            If Cells(i, 1).Value = "DQ" Then
                TotalVolume = TotalVolume + Cells(i, 8).Value
            End If

            If Cells(i, 1).Value = "DQ" And Cells(i - 1, 1).Value <> "DQ" Then
                Startingprice = Cells(i, 6).Value
            End If

            If Cells(i, 1).Value = "DQ" And Cells(i + 1, 1).Value <> "DQ" Then
                Endingprice = Cells(i, 6).Value
            End If
From here told the output this code to go into our new "DQ Analysis" worksheet. This is the basic outline for how we completed the analysis of all stocks as well, but instead of using just DQ, we created an Array of all tickers, and where we had "DQ" in the above code we used our new variable "ticker" instead. This allowed the macro to go through and give is the total volume and percent return for each of the 12 stocks in our original data sheet. The ticker array we used is listed below.

        Dim tickers(12) As String

        tickers(0) = "AY"

        tickers(1) = "CSIQ"

        tickers(2) = "DQ"

        tickers(3) = "ENPH"

        tickers(4) = "FSLR"

        tickers(5) = "HASI"

        tickers(6) = "JKS"

        tickers(7) = "RUN"

        tickers(8) = "SEDG"

        tickers(9) = "SPWR"

        tickers(10) = "TERP"

        tickers(11) = "VSLR"

We also added a timer to our AllStocksAnalysis subroutine, to track how how long it took for the program to run to completion. For the year 2017 the macro took 0.67969 seconds to complete, and for 2018 it took 0.66797 seconds.

Finally, we refactored our code to run more efficiently, which I will describe in the next section.
### Analysis
To refactor our code, we wanted to have our macro only go through the data sheet one time, instead of looping through over and over until all rows were counted. To do this we created variable we called tickerIndex, and set it to 0. Then we created three different output arrays for our volume, starting, and ending price variables. That code is below.

    Dim tickerVolume(12) As Long

    Dim tickerStartingPrices(12) As Single

    Dim tickerEndingPrices(12) As Single

Now that we have those arrays established, we use a similar code as we did in the DQ and All Stock Analysis, only now using our tickerIndex variable to cycle through each of the tickers. That code is below.

    For i = 2 To RowCount

        '3a) Increase volume for current ticker
            tickerVolume(tickerIndex) = tickerVolume(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        'End If
            End If

        '3c) check if the current row is the last row with the selected ticker
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            tickerIndex = tickerIndex + 1
            End If

         'If the next row’s ticker doesn’t match, increase the tickerIndex.
            '3d Increase the tickerIndex.
    Next i

We then had the output displayed in our new worksheet, and formatted that data in clear, decipherable way. Lastly, we had the timer set up for our refactored code as well, and as we can see below this new refactored code was significantly faster and more effecient than our original code.

<img src=Resources/VBA_Challenge_2017.png>
<img src=Resources/VBA_Challenge_2018.png>

## Summary
This project was able to show us one of the major benefits of refactoring code in general. After refactoring our code ran signficantly faster, and was more efficient than our original code. This is the goal when refactoring code, to make code more efficient by executing fewer steps, by using less memory, or by improving the logic/ readability of the code. Some disadvantages of refactoring code could be that you take away some functionality when trying to improve the code, and it can be a time consuming proess to think of new ways to approach a problem. For our VBA code in this project, we clearly saw the benefits of refactoring in code that ran almost 6x faster than the original. Our use of arrays allowed the code to run through our data set once top to bottom, instead of many times. While the original code was very functional with this data set, it could been very slow for a larger data sample. Refactoring our code let us avoid this problem, and now we have better  more functional code we can use for future analysis.
