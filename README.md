# stock-analysis
# Module 2 Challenge - Refactored Stock Analysis Script 
## Overview of Project: 
The purpose of this analysis is to evaluate the efficiency of a code refractor. After creating a workbook using macros to calculate total daily volume and return for multiple stocks, the macro was refactored to make use of arrays to reduce computational complexity and increase overall speed and efficiency. Improved speed is evaluated using a timer function in VBA to calculate the total runtime for the macro. The length of time for the original macro and refactored macro were compared after ensuring that the results for the improved version were identical to the original.  
## Results
### Prior to refractoring 
Reducing the number of times, the program had to loop significantly reduced the amount of time required to run the macro. For example, in the original script, the program looped 12 times, once for each ticker and again for each of the rows in the dataset executing in big O of n^2 time. 

```sh
For i = 0 To 11
ticker = tickers(i)
totalVolume = 0
Worksheets(yearValue).Activate
For j = 2 To RowCount
```
The number of seconds required to run the original script is shown in the screen shot below. 

[![Pre_Ref](https://raw.githubusercontent.com/asanchez116/stock-analysis/master/VBA%20challenge/Resources/Pre_refractor.png)](https://raw.githubusercontent.com/asanchez116/stock-analysis/master/VBA%20challenge/Resources/Pre_refractor.png)

### Refractored 
The refactored macro only looped through the 3013 rows in the dataset once and adjusted the values to the arrays as defined in the If then statements executing in big O of n time. 

[![Post_Ref](https://raw.githubusercontent.com/asanchez116/stock-analysis/master/VBA%20challenge/Resources/Post_refractor.png)](https://raw.githubusercontent.com/asanchez116/stock-analysis/master/VBA%20challenge/Resources/Post_refractor.png)
## Summary: 
### Advantages 
Refractoring code can result in improved design, by focusing on efficiency and implementing data structures to make the code run more efficiently. Data structures like arrays, hashmaps/dictionaries, trees, heaps, graphs, and others can reduce the amount of work needed for the computer to process large amounts of data. It can also help with making the code easier to understand as there are fewer hard coded elements adn confusing loops and often can result in reducing the amount of code needing to be written to acomplish the same task.   

Refractoring can also help with identifying bugs by working through alternative paths that might introduce additional variables that could break the code in unexpected ways. We might also improve readability by adding or improving the comments to provide detailed explanations for how the code works and what to expect from results. 

### Disadvantages
Refractoring can be time consuming and expensive without adding new functionality or utility. There is also the possability of introducing new bugs
### Refractored script 
The refractored module 2 challenge code incorporated arrays to reduce the workload and produce the same outcome 80% faster than the original script. The refactored script was also easier to read as it didnt require keeping track of a nested loop and would result in easier to maintain code when changes/updates are required. 



