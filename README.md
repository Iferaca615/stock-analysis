# stock-analysis
Written with VBA
---

# **Overview of Project:**

- This analysis was created in order to help Steve's parents better visualize, the returns on various green energy stocks during the years 2017 and 2018.   At first, they were specifically interested in mapping the returns and total volume of the stock "DQ", which we did analyze individually in the first module.  However, in the Challenge we were asked to analyze all of the green stocks in the data table to see how they fared in 2017 and 2018.
 
--- 

The table given to us had 3013 rows of data, which to the average eye is a lot of information to read and understand.  

- We created and formatted a table to organize by, stock ticker (name), total daily volume and percent return so that the data would become readable.  All formatting and code was done in the Visual Basic editor.  
- The code allows the user to input the year they wish to see data for, and then outputs a chart as well as a pop up that shows how fast the script was ran.
	
## Refactored code versus Original and why we refactor:
In this challenge we were told to refactor our code in order to run the script faster, and to make the code more concise and effective.  The pop up screens with runTime of the code were used to compare the processing speed of the code before and after refactoring.
- listed below are the files:

	1) [green stocks original file](stock-analysis/green_stocks.xlsm")
	2) [VBA Challenge refactored code](stock-analysis/VBA_Challenge.xlsm)

### Why refactor the code?

Refactoring code has many uses that are extremely important in order to make sure your code is as concise, understandable and effective as possible. During the first run through your code, there are often many instances where code can be changed; to execute the same function, to help the code run better.  In our [refactored code](stock-analysis/VBA_Challenge.xlsm) , there were some changes made that in return, allowed our code to run faster.

1) separating the for loops from nested for loops:
- By separating the for loops into two separate loops, we allowed the code to run continuously from beginning to end instead of looping back within the nested loop and then processing the rest of the code.
 	* ```  RowCount = Cells(Rows.Count, "A").End(xlUp).Row    For i = 0 To 11        ticker = tickers(i)        totalVolume = 0        'loop through rows in data    Worksheets(yearValue).Activate                        For j = 2 To RowCount    'increase totalVolume by the value in the current row                If Cells(j, 1).Value = ticker Then            totalVolume = totalVolume + Cells(j, 8).Value                End If        'set startingPrice for current ticker                If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then        startingPrice = Cells(j, 6).Value                End If            'set endingPrice for current ticker                If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then        endingPrice = Cells(j, 6).Value                End If            Next j                'activate output worksheet        Worksheets("All stocks analysis").ActivateCells(4 + i, 1).Value = tickerCells(4 + i, 2).Value = totalVolumeCells(4 + i, 3).Value = (endingPrice / startingPrice) - 1    Next i
```
		

---
# Data Visualizations:

1) ![VBA_Challenge_2017](resources/VBA_Challenge_2017.png)
2) ![pre-refactor-2017](resources/pre-refactor-2017.png)
3) ![VBA_Challenge_2018](resources/VBA_Challenge_2018.png)
4) ![pre-refactor-2018](resources/pre-refactor-2018.png)
---
# Results:

## analysis of charts:
- In general, green energy stocks in 2017 had far greater returns than in 2018.

- The **highest** performing stock in 2017 was "DQ"
	* "DQ" fell to -62.6% returns in 2018

- The **highest** performing stock in 2018 was "RUN"
	* "RUN" was returning at a percent of 5.5.% in 2017, and increased its returns to 84.0% in 2018
 
