# Excel and Macros in VBA

## Overview of Project
We are tryingIn this project, we are helping Steve, run analysis on all stocks provided in green_stocks.xls sheet using macros. However for this task we are also refactoring the macro that we previously wrote to run analysis on all stocks for sheets "2017" and "2018", such that it takes lesser time to run analysis on all stocks.

### Results

#### Time taken for analysis 2017
<p float="left">
  <img src="Resources/VBA_2017_Pop_Up_Not_Refactored.png" width="250" />
  <img src="Resources/VBA_Challenge_2017_Pop_Up.png" width="250" /> 
</p> 
    

Above are the images of time taken to compelete analysis for 2017 all stocks between original code and refactored code. 
Refactored code takes much lesser time to complete.

Let us look at the difference now for 2018

#### Time taken for analysis 2018
<p float="left">
  <img src="Resources/VBA_2018_Pop_Up_Not_Refactored.png" width="250" />
  <img src="Resources/VBA_Challenge_2018_Pop_Up.png" width="250" /> 
</p>

We see a similar pattern here, refactored code took much lesser time than the original code.

### Summary

#### 1. What are the advantages or disadvantages of refactoring code?
Advantages: Time taken for running our analysis takes less time. Makes it easier to run analysis on more items if needed.
Disadvantages: Refactoring code takes time, if you are not familiar with code, you can often get stuck with code errors.

#### 2. How do these pros and cons apply to refactoring the original VBA script?
Time taken to run our analysis is less. Also we can add more stocks to our analysis and we can get our results in reasonable time.
However when refactoring VBA code, it gets difficult as VBA is not very intuitive. Debugging code is not easy when trying to figure out if changes made to code are working or not.





