# stock-analysis


## Project Overview
This project is to refactor a VBA code which provides stock analysis to a ticker's Total Daily Volume and Overall Return per over each Calendar Year. The refactored code will improve VBA macro's run-time and re-usability/scalability.

## Results

### Refactoring Approach
- Run-time was improved by reducing nested loops and used stored array variables inside of the macro and then writing each array element into the workbook after the analysis is complete. In the original code, the code would write into the excel file for each iteration.
  - In the original code: 
     ```
     For i = 0 To 11
        ticker = tickers(i)
        ...
        For j = 2 to RowCount
        ... 'analysis per "j" ...
        Next j
        ...
        #print results for each i
     Next i
     ```
     - Every row count (j) is iterated through each ticker (range of i), effectively running "i x j" number of rows
   - In the refactored code: 
      ```
      Dim tickerVolumes(12) As Long
      ...
      For tickerIndex = 0 To 11
        tickerVolumes(tickerIndex) = 0
      Next tickerIndex
      
      For i = 2 To RowCount
      ... 'analysis per "i" and store in respective array element tickerVolumes(i)...
      Next i
      
      For i = 0 to 11
      ... 'print tickerVolumes(i)
      next i    
      ```
      - the ticker array length is iterated first, and then the row analysis is iterated afterwards. As such, instead of running "i x j" lines of code as seen in the original code, the refactored code runs "i + j" lines of code (in this case, "tickerIndex + i (2 to row count) + i (0 to 11)", as i is a re-used index variable).
      
      
      
- Usability will be increased by allowing user input to specify the year of analysis, rather than having a hard-coded year in the VBA code.
  - In the original code:
    ```
    Range("A1").Value = "All Stocks (2017)"
    ...
    Worksheets("2017").Activate
    ...
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    ...
    Worksheets("2017").Activate
    For j = 2 To RowCount
    ...
    Next J
    ```
  - In comparison, the Refactored code:
    ```
    yearValue = InputBox("What year would you like to run the analysis on?")
    ...
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    ...
    Worksheets("yearValue").Activate
    ...
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    ...
    Worksheets("yearValue").Activate
    For j = 2 To RowCount
    ...
    Next J
    ```
### Run Times
- Run time and results of refactored code:
  - ![2017 Refactored Stock Analysis](resouces/VBA_Challenge_2018.png) ![2018 Refactored Stock Analysis](resouces/VBA_Challenge_2018.png)<br/> 
  - Total Run Time: 
     - 2017: 0.1836 seconds  <br/>
     - 2018: 0.1875 seconds
- In comparison, the run time and results of original code:
  - ![2017 Original Stock Anlaysis](resouces/Original_VBA_Challenge_2017.png) ![2018 Original Stock Anlaysis](resouces/Original_VBA_Challenge_2018.png)<br/> 
  - Total Run Time: 
     - 2017: 0.9727 seconds <br/>
     - 2018: 0.9805 seconds <br/>

## Summary
- In general, refactoring code has the following
  - Advantages:
    - Maintainability and improvement to existing code. By definition, refactoring is taking an existing code/file and simply improving it - whether it means to improve extensibility, usability, or run-time of the code, the developer/anaylst does not need to redefine the problem, and should not need to redefine its output nor its dependancies.
  - Disadvantages:
    - The quality of the refactored code is largely dependent on its original code. If the original code is poorly documented, refactoring the code may take much longer than necesary. In other cases, if the base logic of the code is constructed ineffectively, there may be times where the developer/analyst would rather build the code afresh.
- In both the refactored and original script of this specific VBA project
  - Pros:
    - The script is great for investors who are interested in tracking only a specific set of stocks in their portfolio. In some cases, investors may diversify into many different stocks but those stocks may not neccesarily be a part of their "core" portfolio they wish to analyze.
  - Cons:
    - The tickers are hardcoded into the script, and not only are the numbers of tickers analyzed limited, but the code also will not readily anaylze any new tickers.
