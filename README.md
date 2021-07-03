# stock-analysis

## Overview of Project
This workbook contains a series of stocks that were traded during 2017 and 2018. A client came forward and asked if an analysis can be done on these stocks to determine which there parents should invest for their retirement. Thus, I utilized VBA and wrote a code to get back the average return for a given year. In the first file ([green_stocks](https://github.com/bazinga183/stock-analysis/blob/main/green_stocks.xlsm)), I utilized a code that got the same reaults, but was running at a pace that was inferior to what the clients desired. 
Therefore, in the most up-to-date file ([VBA_Challenge](https://github.com/bazinga183/stock-analysis/blob/main/VBA_Challenge.xlsm)), I refactored the VBA code to incorporate a system of arrays so that the code would run at a fracture of the time.

## Results

### Comparing the Stock Returns for 2017 and 2018
If we look at 2017, we can notice that almost all stocks yielded a positive return with the exception of TERP. On its surface, this means that the client cannot go wrong unless they choose the only unprofitable stock:  TERP.

![2017Stocks](https://user-images.githubusercontent.com/46951897/124365299-a91e8500-dc0c-11eb-9f40-9d3fbae391af.PNG)

However, upon further analysis of 2018, the clients were alarmed to find that almost all stocks had a negative return, except for ENPH and RUN.

![2018Stocks](https://user-images.githubusercontent.com/46951897/124365375-3a8df700-dc0d-11eb-92df-3a5fb9ce2210.PNG)

Despite both stocks having a positive return in 2017, ENPH appears to be the superior choice over RUN because ENPH returned 129.5% and 81.9% for 2017 and 2018, respectively. This outpaces RUN's overall returns of 5.5% and 84%, therefore, the clients can choose ENPH as the safer choice.

### Comparing the Original and New Times
The original VBA code in ([green stocks](https://github.com/bazinga183/stock-analysis/blob/main/green_stocks.xlsm)) yielding times of 0.9335938 and 0.90625 seconds for 2017 and 2018, respectively.

![2017(1)](https://user-images.githubusercontent.com/46951897/124365425-8771cd80-dc0d-11eb-9af4-9fc4eb40fa59.PNG)
![2018(1)](https://user-images.githubusercontent.com/46951897/124365441-a7a18c80-dc0d-11eb-81ad-160a1d1cc0c7.PNG)

The new code yielded faster run times of 0.1484375 and 0.1328125 seconds for 2017 and 2018, respectively.

![2017(2)](https://user-images.githubusercontent.com/46951897/124365448-bdaf4d00-dc0d-11eb-864d-adfca8ce7893.PNG)
![2018(2)](https://user-images.githubusercontent.com/46951897/124365451-c011a700-dc0d-11eb-9e3d-3ac7fddf5961.PNG)

Running a comparison of the original and refactored codes yields a considerable inprovement in total run time by apoximately .8 seconds, or 85%.  

This was in part because of the implementation of arrays for the projects:

```
    '1a) Create a ticker Index
    
    tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
```
The refactored code also incorporated these arrays as shown:

```
 For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            '3d Increase the tickerIndex.
            
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
```

## Summary

### Advantages and Disadvantages of Refactoring VBA Code
Refactoring code can be very beneficial for larger datasets where the amount of data for a simple code can be overwhelming. It can also aid in going back over previous code and correcting or simplifying prior mistakes which make the code read better and execute at a faster pace. However, the fault in this endeavor is that it can be time-consuming for the programmers and clients alike if results are needed at an expedient point in time.

### Advantages and Disadvantages of Refactoring This Code
Whenever the code was refactored to use arrays, the advantage was that the times were cut down drastically. If this data set were to add a larger share of the stock market, then this refactored code would save future clients huge amounts of time in the long term. As it stands, the main disadvantage of writing this refactored code is that the change in time is very minimal by human standards; less than a second of total difference. The original code fulfilled the same purpose and is perhaps better utilized for the smaller dataset I was presented.
