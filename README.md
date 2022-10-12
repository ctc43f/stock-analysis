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
In the refactored code, we only scanned through the 3,000 rows only one time.  As we went down the rows of data, we did three things:
1. Kept a running total of volume as long as the stock ticker label did not change.
2. Kept track of the price we saw the first time we encountered the new stock ticker label.
3. Looked to see if the next row of data was for a different ticker and, if it was, grabbed the end price and made a record of the cumulative volume before initializing our running total variable.

Effectively we removed the portion of the original code regarding j = 2 to RowCount to reduce overall processing time.

## Summary

### - What are the advantages or disadvantages of refactoring code?
Advantages of refactoring code include:
- Reductions in processing time and memory usage; the latter is important because the processing happens on a local machine and poorly written code can fill RAM and cause the program to crash.  An example would be attempting to make multiple copies of a data-heavy tab as a processing step or attempting to create a complex array function in Excel that prevents the macro from completing.
- Ability to separate and distribute the work across multiple individuals.  If the code is highly compartmentalized and modular, you could assign individuals to each refactor a portion and then run tests on the entire macro as portions are complete.  This is much easier when subroutines are contained in their own macros and then the high-level macro calls those subroutine macros; you can refactor and just ensure the outputs are the same while not breaking other portions of the main macro.

Disadvantages of refactoring code include:
- Value of refactored code vs. benefit of time spent: if the macro will be used sparingly or process time and memory are not factors, the time spent to refactor may not yield the same value.  If you pay ten high-cost resources for six months to refactor and you save 1 second of processing time, it likely wasn't worth the effort.
- Complexity of the code: if the code has several layers of nested functions, is not adequately documented via comment, or things like variable declarations are scattered through the code block, it might be hard or impossible to refactor and obtain the same results because of some embedded dependency.
- Refactoring areas that don't generate efficiencies.  If the code is complex and you are having issues with memory or processing time, it might not be clear at first glance understanding what portion of the process is the root cause of the issues.  Refactoring portions of code that don't generate benefit is wasted time.

### - How do these pros and cons apply to refactoring the original VBA script?
