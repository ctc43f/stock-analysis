# Analyzing Stock Prices

## Overview of Project
Our client asked us to provide an automated way to analyze and provide summary statistics for stock prices in 2018.  I created a macro that provided a summary table of the handful of stocks; however, the client's concern is that the created macro will not perform well when applied to a very large data set.

**The goal of this project is to refactor the existing VBA code to determine if it will improve performance, and to then explain why.**

## Results
By refactoring the code, I was able to improve the total processing time by 50%.

**Original 2018 Processing Time**

![Image 1: Original Processing Time](/Resources/VBA_Challenge_2018_Original.PNG)

**Refactored 2018 Processing Time**

![Image 2: Refactored Processing Time](/Resources/VBA_Challenge_2018_Redone.PNG)

The main driver of the longer processing time in the original set of code was that we iterated through all 3,000 rows of data 12 times, once for each of the individual stocks.  In essence, we started from the beginning for each individual stock even though we knew its volume and price information wasn't in the rows we scanned during the last iteration.

```
   For i = 0 to 11
      (Initialize values at 0)
      For j = 2 to RowCount
      (Iterate until you see the first row of stock in question and get volume and start/end pricing)
      Next j
   Next i
```
In the refactored code, we only scanned through the 3,000 rows one time.  As we went down the rows of data, we did three things:
1. Kept a running total of volume as long as the stock ticker label did not change.
2. Kept track of the price we saw the first time we encountered the new stock ticker label.
3. Looked to see if the next row of data was for a different ticker and, if it was, grabbed the end price and made a record of the cumulative volume before initializing our running total variable.

### - What are two conclusions you can draw about the Outcomes based on Launch Date?

### - What can you conclude about the Outcomes based on Goals?

### - What are some limitations of this dataset?

### - What are some other possible tables and/or graphs that we could create?
