# VBA of Wall Street

## Overview of Project

### Purpose
	The purpose of the assignment was to refactor the original AllStocksAnalysis script to create a script that generated the analysis of
the 2017 and 2018 tickers for Return and Total Daily Volume with a faster and more efficient process. The assignment gave students the opportunity
to understand the code from an editing perspective and see how the same processes can be conducted with different code. It led to a greater 
understanding of the process and knowledge of different ways to code.

## Results

### Analysis of AllStocksAnalysisRefactored()
	This assignment showed the students they can reconsider using for loops, in particular nesting for loops, and suggests that the less nesting loops there are the faster the program will
execute. The 2017 data was initially run at 1.926 seconds and was reduced to 0.371 seconds, an 80% decrease in runtime. The 2018 data was initially run at 1.711 seconds and was reduced to 0.227
seconds, an 86% decrease in runtime. The main reason for these significant reductions in time is the removal of a nested for loop in the original code. 

## Challenges and Summary

### Challenges and Difficulties Encountered
	I received multiple overflow errors. The first was based on the Volume variable and was resolved when I changed the array from As Long
to as Double. The second was based on the Return variable. I was dividing by zero because the tickerStartingPrice and tickerEndingPrice were
not updating properly and needed to be reconfigured. This was resolved by editing the conditional in the If-Then statement from "tickerIndex"
to "tickers(tickerIndex)" to properly select the ticker string values instead of an integer.
 
### Summary
	In general refactoring code can lead to a greater understanding of the processes in the code for the programmer and hopefully lead to 
faster runtimes and efficiency of the code. The programmer will be able to use the lessons from each time the code is refactored in subsequent 
subroutines and macros. The main disadvantage is that this procedure takes time that could be spent on other various projects the programmer
may have. An additional advantage is that the programmer will be more likely to find any errors are areas for improvement for the program.
	Refactoring the VBA challenge script saw the advantage of a decreased runtime of over 80% for the given datasets. As the programmer,
I better understood how to index arrays and reference values in cells from worksheets. The disadvantage was the additional time it took 
to review and edit the code, but this is a small price to pay for learning and optimization.   

