# **VBA Challenge: Stock-Analysis**

## Project Overview
Utilizing VBA code, the project provided the ability to automatically aggregate the volume and yearly return from a dataset of 'green stocks' in order to visualize the best stocks to invest in based on historical performance. 
The focus of the VBA Challenge was specifically an exercise in refactoring. The purpose of refactoring is to take existing code functionality and configure the code steps and logic to achieve higher performance. 
 
## **Analysis/Results**

**Original stock analysis**: 
This subroutine provided the correct output for ticker name, volume and return for a small dataset of green stocks. The code for obtaining the volume and return of 12 tickers utilized two nested loops. 
The inner loop stepped through all the rows in the respective 2017 and 2018 dataset worksheets before the outer loop associated the details to the respective rows and columns. Then the output worksheet (re: Stock Analysis) was formatted for look and feel. 
From start to finish the code in the 2018 original sample ran in 2.27 seconds. Initially this automated return may seem fast, but it is a very small subset of 12 tickers and requested output (3 fields). 
Investment decisions require quick turnaround analysis. If a client wanted to see an unfiltered list of stock exchange tickers (ex: NYSE) and required additional output for investment decisions, a nested loop may not be optimal. 

**Refactored stock analysis**: 
This subroutine **also** provided the correct output for ticker name, volume and return for a small dataset of green stocks. The goal of this subroutine was to take the original code and refactor the statements in order to improve performance.  
This was obtained by identifying a ticker index and associating it to several arrays. The arrays ticker, volume, and return must share the same datatype, respectively. The advantage to using an index and arrays was the elimination of the nested loop, and reduction of memory and steps to associate the stock detailed inputs to the aggregated outputs.   
Then similar to the original stock analysis, the output worksheet (re: Stock Analysis) was formatted for look and feel. 
From start to finish the code in the 2018 refactored sample ran in .46. This is a notable difference from the original code. 


### VBA Challenge: Original vs. Refactored Performance Comparison
![Performance Original 2017]()
![Performance Refactored 2017]()
![Performance Original 2018]()
![Performance Refactored 2018]()

### VBA Challenge: Original vs. Refactored Code Comparison 
![Refactored Code Support](https://github.com/ljlodl5/Stock-Analysis/blob/main/Resources/VBA%20Challenge%20Code%20Comparison.png)

## **Results**
### **VBA Script (Stock-Analysis) Advantages and Disadvantages of Refactoring Code**
#### Advantages
The most noticeable advantage to the code refactoring is the increased performance. The factored stock-analysis macro runs ~5x faster than the original. This can be attributed to the index/array advantage and the reduction of the nested loop. 
Also when running the macro multiple times the variation in run speed is greater when the code underperforms. The refactored code provides a stronger predictor of when it will finish than the original code. 
Begin/end variance may not seem like a large impact for a small dataset, however a robust dataset and strict deadlines may require stronger reliability on assessing when the analysis will be complete. 
Original Code Run Times:   (Highest=5.5 seconds; Lowest=2.2 seconds)
Refactored Code Run Times: (Highest=.88 seconds; lowest = .3) 

#### Disadvantages
There are a couple disadvantages to refactoring code specific to this exercise. For one, it took additional time to reconfigure an already working concept with no guarantee of performance improvement.     
Second and perhaps minor by comparison, it is much easier to conceptually understand the routines of a nested function than to associate arrays with an index function. 
I encountered a small issue of my own as I erroneously rebuilt a nested function and found my 'refactored' performance was virtually the same as the the original code.*
*Credit to https://github.com/AndrejaCH/stock-analysis/blob/master/Graphics/CodeRefactored.PNG for identifying a specific placement of 'For Loop' missing in my version of the refactored code.*
    


### **General Advantages and Disadvantages of Refactoring Code**
#### Advantages 
There are multiple advantages to refactoring in general. Refactoring not only can improve performance by simplifying steps, logic, and reducing memory it can also just simply clean up the code for future users. 
It is also important to note the original needs of a client can change over time. Data inputs and outputs can be dynamic and grow; the same manner of coding that worked originally can become outdated and drive down performance. 
#### Disadvantages
General disadvantages to refactoring are time and project risk. Not only is a developer coding and unit testing, regression and parallel QA/UAT time may also be required to ensure key functionality/data was not changed in error. 
If change in function/data does occur, additional time may be spent trying to uncover a version of truth. 
There are no guarantees that the perfomance benefit will be worth the time invested. It is important to gauge the level of service/time you are trying to achieve and the amount of time/$ you are willing to invest to refactor. 




