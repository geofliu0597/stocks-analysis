# Module 1 Challenge - Analysis of Kickstarter Data through Excel

## Overview of Project

The purpose of refactoring the code was to make the code run faster so that in the future it could potentially be applied to a dataset that includes thousands if not more tickers of data. The ability to make the code run faster is important so that we can properly make the analysis on larger and larger sets of data without increasing the run-tiem cost by too much.

## Results

#### Stock Performance
![](/resources/stock_performance_2017.png)
![](/resources/stock_performance_2018.png)

As is already shown on the Module 2 Challenge page, the stock performance in 2017 was significantly better than the performance of 2018. This is to be expected and refactoring the code had nothing to do witht he performance of the stocks themselves. It was primarily used as a measure to check whether or not the code was working as intended. 

#### Script runtime
![VBA Challenge 2017](/resources/VBA_Challenge_2017.png)
![VBA Challenge 2018](/resources/VBA_Challenge_2018.png)

The results of refactoring the code were amazing. The previous runs took approximately 0.67 seconds whereas now they take approximately 0.07 seconds, reducing the run-time of the code by almost 90%.

## Summary

#### Advantages

The clear advantage of refactoring the code in this way is to significantly reduce the runtime of the code, allowing the same code to be used on massive data sets without wasting time just waiting for the code to run. In this case, larger data sets would make the original code take exponentially longer to run as there would be more data to run in each iteration, as well as more iterations to go through with more tickers to keep track of.

#### Disadvantages

The primary disadvantage i can think of is the extra memory required by the code itself, since we keep all this information within data arrays and then dump them all onto the spreadsheet at the end, rather than dumping the information one ticker at a time. With modern data limitations though I can't really see this becoming a problem even with larger data sets with the small amount of calculations we are doing.

Another disadvantage is just the time it takes to refactor the code in this way. It is always a cost-benefit analysis of whether or not the resources taken to improve the code will pay dividends in the future. In this case, it took me an hour or two to refactor this code, but I am only saving 0.5 seconds each time. With this particular data set there isn't much reason to run it many times a day since this is data pertaining to a full year and grouped in that way. It will take centuries if not millenia for the time to be regained. On the other hand if I am saving minutes or even hours for significantly larger datasets then it'll pay dividends to refactor the code.



