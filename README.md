# Analyzing Green Energy Stocks

## Overview of Project
- The purpose of this project was to create a VBA Macro that automates the calculation of _yearly trade volume_ and _yearly return_ of 12 green stocks. The macro is able to quickly anaylyze thousands of data points as well as uncover trends that help users make financial decisions. 

## Results

**Stock Performance Comparison Between 2017 and 2018**
- Using our refactored code, we calculated the stock performance for each stock. With the exception of ENPH and RUN, green stocks fared worse in 2018 compared to 2017.

<img width="754" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/10199828/123672692-a0ddd880-d80d-11eb-834a-8746bf21fc1d.png">
<img width="755" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/10199828/123672696-a20f0580-d80d-11eb-8d65-8ce6755ae45c.png">


**Measuring VBA Code Performance Between Original and Refactored Code**
- Two types of code were written to perform stock analysis: one that used two for loops, and a refactored code that used an index and an array. Below is each macro's run time (in seconds) for 2017 and 2018 stock data. The VBA excel macro can be viewed [here](https://github.com/gabalita/stock-analysis/blob/main/VBA_Challenge.xlsm).

<img width="427" alt="Performance Data" src="https://user-images.githubusercontent.com/10199828/123674504-acca9a00-d80f-11eb-8885-f188aa73f6cd.png">



## Summary

**Advantages and disadvantages of refactoring code in general**
- Overall, the advantages of refactoring code are that it makes the code cleaner, simpler, easier to understand, and more efficient. While the benefits of refactored code might not be always noticeable in the short term, over time it helps users identify bugs and substantially increases the scalability of the code. The disadvantages of refactoring code are that it requires additional investments in time and money. 

**Advantages and disadvantages of the original and refactored VBA script**
- Although refactoring the VBA macro took a couple hours, our refactored script is easier for future coders to review and understand. It also significantly increases the efficiency of the code by lowering the performance times from a couple mintues to nearly a tenth of a second. 

