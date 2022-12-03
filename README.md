# Stock-analysis
Module 2 VBA Challenge
## Overview of Project
In this project our goal was to provide stock performance information on other green energy stocks for Steve's parents.  Up til now, his parents have been investing primarily in DAQO and Steve believes they should diversify across other green energy companies.  Our analysis will help them to see performance of other green energy companies so they can see how it compares with their current investment with the hopes that they will diversify their investments.
## Results
### Comparison of Stock Performance
In 2017, all but one of the green enery stocks shows an increas in performance.<img width="249" alt="Stock_Performance_2017" src="https://user-images.githubusercontent.com/56225030/205451127-ebb8fe5b-aede-49b5-9037-b8821c1c0f74.png">  In fact, the DAQO company, stocker ticker DQ, had the best performance with 199.4% return on the stock price. In 2018 most of the stocks fell in price, except ENPH and RUN.<img width="250" alt="Stock_Performance_2018" src="https://user-images.githubusercontent.com/56225030/205451691-70e36b3c-630a-4051-9c2c-d8049d052d35.png">. Although DQ did not completely return to the starting price in 2017, it did lose about 62% of its value.  Again, this is another good reason for Steve's parents to diversify their portfolio.  
I would highly recommend looking at the ENPH stock.  Although the stock price is relatively low comared with DQ, the stock increased in 2017 and again in 2017, for another 81%.  Again, the price is lower which means that they could invest at a higher volume.  

### Comparison of Code Performance
The original VB code that primarily used a nested loop to summarize volumes and return on price ran at approximately 36,000 cpu seconds for both years; whereas, the refactored code only took .09 and .06 cpu seconds for 2017 ![alt text](https://github.com/ericajackson8/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png?raw=false) and 2018 ![alt text](https://github.com/ericajackson8/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png?raw=true), respectively.

## Summary
Refactoring code is the process of making modifications to the code without altering the functionality.  Modifications may include streamlining the logic, order, and streamling or modifying syntax. Again, all of this is done without altering the output or functionality.
### Advantages and Disadvantages of Refactoring Code
Some advantages of refactoring code are:
1. Providing for a more streamlined appearance, and/order of the code to make the code easier to understand.  Also providing for more effective documentation.
2. Improving runtime performance by streamling the syntax to optimize cpu processing time.
3. Improve capacity to debug and troubleshoot code.

Although disadvanteges are fewer, there are some.
1. Re-inventing the wheel.  You basically already completed the code and it works.  If it aint broke, don't fix it.
2. Possibility modifying the functionality.  Always make sure the original code is saved so you can refer back.  
3. There may not be more efficient syntax then what you are currently using.  In this case, make sure what you are using is organized well, and you review for any use of magic numbers and other hard coded references.

### How do the Pros and Cons Apply to this Project
Refactoring of the VBA code for the Stock analysis project made the code easier to follow by removed the complexity of a nested loop. The use of arrays allowed us to create, store, and provide explicit references to the values for ticker, daily volume, and return on price for each company. This made debugging easier because there was less code, and we could more readily identify what was happening to each specific ticker at various points in the code.
