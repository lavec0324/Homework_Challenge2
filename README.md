# Homework_Challenge2
Homework Challenge 2
## VBA Refactoring

### Overview of Project

This project is intended to analyze stock data for 12 stocks and for multiple years.  The specific intent of this exercise is also to use arrays to speed up code that was created previously.

### Stock performance

Through the analysis we could see that stock performance overall was better in 2017 then it was in 2018. 

### Execution times

Performance times were greatly enhanced from code that did not contain arrays.  Previously the code was running at ~ .67 seconds for both years.  Currently we see an increase performance for both years as can be seen here:

![2017 performance](https://github.com/lavec0324/Homework_Challenge2/blob/main/Resources/VBA_Challenge_2017.PNG)

![2018 performance](https://github.com/lavec0324/Homework_Challenge2/blob/main/Resources/VBA_Challenge_2018.PNG)

Code that help produce these efficiencies include:

```
        1b) Create three output arrays
        
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
    
```

and

```

    4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For m = 0 To 11
        
        Worksheets("AllStocksAnalysis").Activate
        Cells(m + 4, 1).Value = tickers(m)
        Cells(m + 4, 2).Value = tickerVolumes(m)
        Cells(m + 4, 3).Value = tickerEndingPrices(m) / tickerStartingPrices(m) - 1
     
    Next m   
```

### Summary
#### What are the advantages or disadvantages of refactoring code?

The clear advantage was the enhanced performance to run time and clear code that could be utilized in 
the future.  For larger scale projects these run times also equate to money so that would allow companies to save money with more efficient code.  The only disadvantage that I
could see was the additional time needed to refactor the code but if it was used repeatedly this would payoff from the upfront investment.

#### How do these pros and cons apply to refactoring the original VBA script.
The pros and cons are similar to the general pros and cons.  The upfront time and investment needed to refactor the code was there (especially for a novice. :smile:
as a disadvantage.  The advantage was the efficiency gains from refactoring and learning how to use arrays! 🥳

