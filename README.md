# Stock-Analysis
## Overview
* Steve's parents are passionate about green energy and want to invest their money in DAQO New Energy Corp. However, Steve is worried about the diversifiaction of his parent's funds and wants to analyze the performance of other green energy stocks. In order to do this he has created an excel file containing the stock data he wants us to analyze. Through VBA we will write a code for Steve to perform the analysis in qiuck and accurate manner.
--
## Results
 * In our analysis we compared the performance of 12 companies in the years 2017 & 2018.
 * The metrics we used to measure were Total Daily Volume and Return
 * To measure Total Daily Volume & Return for each year we wrote the code below:
 ``` 
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8)
    
    If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
     End If
  
    'If  Then
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
        tickerIndex = tickerIndex + 1
 ```
 ```
 For i = 0 To 11
    
   Worksheets("All Stocks Analysis").Activate
   Cells(20 + i, 1).Value = tickers(i)
   Cells(20 + i, 2).Value = tickerVolumes(i)
   Cells(20 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
 ```
 * To get total daily volume the code summed up the daily volume for each stock individually.
 * To get the return for the year the code found the starting price & ending price on the first & last trading day of the year. Once it got those values it divided ending price by starting price and subtracted that value by 1 to get the return for the year.
 
 
![image](https://user-images.githubusercontent.com/67936161/88487775-e24e6f80-cf3c-11ea-859d-e5cf093ee6d8.png)
![image](https://user-images.githubusercontent.com/67936161/88487787-0447f200-cf3d-11ea-804b-667fa25eda43.png)
* **2017**: All stocks except TERP had a positive return for the year. DQ had the highest rate of return at 199.4%. However, it had the lowest amount volume. What this means is that although DQ had really good performance in regards to their return for the year, the low amount of volume indicates that not a lot of investors were confident in trading this stock as compared to others in the same industry. The low volume could indicate trading volatility or market uncertainty and thus we cannot deem DQ a good investment soley based ont the rate of return.
* **2018**:
