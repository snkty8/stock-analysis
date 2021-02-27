# Stock Analysis With the Use of VBA
Using VBA and Excel Analyze Stock Data


1. Overview of Project: Explain the purpose of this analysis.

    The purpose of this analysis is to refactor the code.  In turn, making it more
    efficient.  As stated in the challenge, the code works well for a dozen stocks, but it might not work as well for thousands of stocks.  Steven would need the code to 
    work faster and more efficient so he can run evalute more stocks for his 
    parents.

2. Results: Using images and examples of your code, compare the stock performance       between 2017 and 2018, as well as the execution times of the original script and the refactored script.

    Looking at the results from 2017 and 2018, 2017 was definitely a more profitable year for stocks.

    ![image](https://github.com/snkty8/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

    ![image](https://github.com/snkty8/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

    From a coding perspective, incorporating the tickerIndex and output arrays (tickerVolumes, tickerStartingPrices, and tickerEndingPrices) did allow the code to move much faster.  The orginally code from green_stocks had a slower read time, as seen in the pictures below compared to the pictures above.

    ![image](https://github.com/snkty8/stock-analysis/blob/main/Resources/Green_Stocks_2017.png)]

    ![image](https://github.com/snkty8/stock-analysis/blob/main/Resources/Green_Stocks_2018.png)


    In the refacted code, but 2017 and 2018 ran in 0.15625.  This was a very big improvement from 0.765625 for 2017 and 0.734375 for 2018.  

3. Summary: In a summary statement, address the following questions.
    1. What are the advantages or disadvantages of refactoring code?

    The biggest advantage to refacting the code was that it made the run time faster.
    Therefore making the code more efficient.  The only disadvantage was the code got a little longer versus the original code. But, it is a nice trade off to make the code move more efficiently.

    2. How do these pros and cons apply to refactoring the original VBA script?

     Indexing was a huge part of refactoring the data.  In the original data, loops were created and abitriaully looking for certain values. In the refracted data, the index created a specific point for each ticker to follow.



