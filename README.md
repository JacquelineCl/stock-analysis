# stock-analysis
module 2

## Project Overview
This project aims to refactor the original workbook's code so that it runs faster and could be used to analyze the entire stock market in the future.

## Results
After trying multiple ways, the code was eventually adjusted to use a tickerIndex and 3 arrays to calculate the <i>Total Daily Volume</i> and <i>Return</i> for each stock listed in the sheets. 

To fully enable this workbook to be used on all the stocks in the stockmarket, code will need to be added that automatically fills tickers() with the stock tickers contained in the active sheet which is outside the scope of this project. 

Refactoring the code resulted in the code running faster for 2018 stock data. 
Before:
![2018_Before](https://github.com/JacquelineCl/stock-analysis/blob/e27c2896d68add6a4a4591996493295227fd1d89/Resources/2018Before.png)
After:
![2018_After](https://github.com/JacquelineCl/stock-analysis/blob/e27c2896d68add6a4a4591996493295227fd1d89/Resources/VBA_Challenge_2018.png)

2017 stock data also ran faster. 
Before:
![2017_Before](https://github.com/JacquelineCl/stock-analysis/blob/e27c2896d68add6a4a4591996493295227fd1d89/Resources/2017Before.png)
After:
![2017_After](https://github.com/JacquelineCl/stock-analysis/blob/e27c2896d68add6a4a4591996493295227fd1d89/Resources/VBA_Challenge_2017.png)

## Summary
Overall, refactoring code can result in code that is easier to read, easier to maintain, and runs faster. There is a risk that the updated code could produce outputs that are different from the current outputs. 

The advantage of the original VBA script for this project is that it was less complicated and easier to understand. It also took longer to run than the refactored code and could not scale up to 