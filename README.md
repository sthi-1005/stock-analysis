# stock-analysis


## Project Overview
This project is to refactor a VBA code which provides stock analysis to a ticker's Total Daily Volume and Overall Return per over each Calendar Year. The refactored code will improve VBA macro's run-time and re-usability/scalability.

## Results

### Refactoring Methods
- Run-time was improved through the use of stored array variables inside of the macro and then writing each array element into the workbook after the analysis is complete. In the original code, the code would write into the excel file for each iteration.
- Usability will be increased by allowing user input to specify the year of analysis, rather than having a hard-coded year in the VBA code.

### Run Times
- Run time and results of refactored code:
  - ![2017 Refactored Stock Analysis](resouces/VBA_Challenge_2018.png) ![2018 Refactored Stock Analysis](resouces/VBA_Challenge_2018.png)<br/> 
  - Total Run Time: 
     - 2017: 0.1836 seconds  <br/>
     - 2018: 0.1875 seconds
- Run time and results of original code:
  - ![2017 Original Stock Anlaysis](resouces/Original_VBA_Challenge_2017.png) ![2018 Original Stock Anlaysis](resouces/Original_VBA_Challenge_2018.png)<br/> 
  - Total Run Time: 
     - 2017: 0.9727 seconds <br/>
     - 2018: 0.9805 seconds <br/>

## Summary
- There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
- There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).
