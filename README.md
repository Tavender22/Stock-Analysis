# Stock-Analysis Challenge Summary Report
 
 ## Project Overview
 Steve, an friend helping his parents with their investment portfolio, asked for help in developing a model to analyze the returns for various grean stocks. As part of the analysis, a VBA macro was developed to review the returns and volumes for all stocks.

### Improving the Code
Following the initial VBA code development, which took much too long to run, it was decided to try to refactor the code to speed the macro run-times.

## Analysis of the VBA Code
### Original Code and Run-Times
The original code was successful in achieving the desired outputs. However, with run-times for 2017 and 2018 over 70 second each, the macro was not performing in an expected efficient way.
![Original 2017 Runtime](https://github.com/Tavender22/Stock-Analysis/blob/master/VBA%20Mod%202%20Challenge/Resources/2017%20-%20Original.png)
![Original 2018 Runtime](https://github.com/Tavender22/Stock-Analysis/blob/master/VBA%20Mod%202%20Challenge/Resources/2018%20-%20Orignal.png)

### Challenges of the Original Code
The original code was written to run through the database multiple times in its entirety:
 For i = 0 To 11
                'set starting values to 0
                ticker = tickers(i)
                TotalVolume = 0
                
                For j = 2 To rowEnd...
With both For loops running, and over 3,000 lines of data, the code was taking a significant amount of time to run and produce its outputs. The challenge in refactoring the code was to reduce the number of times the code ran over the entire spreedsheet. 

### Refactored Code and Run Times: Improving Results
With the objective of improving the run times, and ultimately seeing the macro run in a reasonable amount of time, it was necessary to reconsider the dual For loops.

This was accomplished by setting up an array representing the 12 tickers, as well as each data point:

    '3) Initialize array of all tickers
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
    ...
     '5b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerstartingPrice(12) As Single
    Dim tickerendingPrice(12) As Single
    
    By creating 12 data points tied to the tickers, we are able to run each If..Then statement while working our way through the data once, vs. re-running it 12 times.
    
    Once all of the required calculations are made, we then run the output code, again streamlining the information flow.

## Conclusions and Results
By reducing the number of iterations throught the data worksheet, we are able to significantly improve the run-times for the refactored code.
![Refactored Run-Time 2017](https://github.com/Tavender22/Stock-Analysis/blob/master/VBA%20Mod%202%20Challenge/Resources/2017%20-%20Refactored.png)
![Refactored Run-Time 2018](https://github.com/Tavender22/Stock-Analysis/blob/master/VBA%20Mod%202%20Challenge/Resources/2018%20-%20Refactored.png)

Overall the more efficient process of indexing the calculations, and reducing the number of times the data set needed to be repeated, we were able to drastically reduce run-times, making the macro much more effcient for Steve's purposes.

    
