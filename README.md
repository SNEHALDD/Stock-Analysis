# Refactor VBA Code

## Overview of Project
This project expects me to refactor the existing code which includes factors such as taking less time to run, making this code more smarter using for loops, using less memory, improving logic.

### Purpose

Steve wants to analyze data of all entire stock market of all years and find out where his parents can invest to get better returns. 

## Analysis and Challenges

Deliverable 1: Refactor VBA code and measure performance
*Challenging part of this analysis was assigning and declaring proper variables and numbers avoiding magic numbers

Deliverable 2: Write analysis

### Analysis of Deliverable 1

*  Created a ticker Index
*  Created three output arrays    
*  Created a for loop to initialize the tickerVolumes to zero. 
*  Loop over all the rows in the spreadsheet. 
*  Increased volume for current ticker
*  Checkd if the current row is the first row with the selected tickerIndex.
*  Checked if the current row is the last row with the selected ticker
*  Created Loops through the arrays to output the Ticker, Total Daily Volume, and Return.

### Challenges and Difficulties Encountered

* Magic numbers created multiple issues which affected the output of the code
* Not declaring values and creating MsgBox at wrong places resulted in iterating MsgBox for 3000 times.

## Results

###Links of VBA_Challenge_2017.png and VBA_Challenge_2018.png
2017 - https://github.com/SNEHALDD/Stock-Analysis/blob/e5af56949b3ad35584289cfaa89d7c37628205c4/Resources/VBA_Challenge_2017.png
2018 - https://github.com/SNEHALDD/Stock-Analysis/blob/e5af56949b3ad35584289cfaa89d7c37628205c4/Resources/VBA_Challenge_2018.png

Refactored Code ran in 0.07421875 seconds for 2017
Refactored Code ran in 0.08984375 seconds for 2018
Execution time of original script - 7.707031 seconds 

Year 2017 had better returns as compared to 2018 year


##Summary
1) What are the advantages and disadvantages of refactoring code?
Ans - Advantages of refactoring - Run time decreases, we do not need nested for loops in refactoring code, code is using less memory, logic of the code improved
Disadvantages of refactoring- Coding became challenging and prone to make more mistakes.

2) How do these pros and cons apply to refactoring the original VBA script?
Ans - Refactored code is more efficient and takes less Time. If I still keep using original VBA script, performance of output will degrade. I tried to apply better logic to decrease run time.

