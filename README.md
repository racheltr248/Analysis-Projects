# VBA-Challenge

## Overview of Project
My client asked me to create a way to analyze thousands of stocks from multiple years using Microsoft Excel and VBA in order to determine the best investment options. 

## Results
### Original Analysis
The first macro I created analyzed the Total Daily Volume and Return for twelve potential stocks. 

![Original analysis](/Resources/Original_analysis.png) 

While this macro did what was needed, it was a bit slow. 

![Speed of original macro for 2017](/Resources/VBA_2017_original.PNG)
![Speed of original macro for 2018](/Resources/VBA_2018_original.PNG)

My client was interested in making the process more efficient so he could use it for future analyses. 

### Refactored Analysis
In order to run the analysis more efficiently for my client, I refactored the original code so it iterated over all of the data for a given year once as opposed to multiple times.

![Refactored analysis](/Resources/Refactored_analysis.png)

This reduced the run time by 84.17% for 2017 and 83.62% for 2018, a significant improvement.  

![Speed of refactored macro for 2017](/Resources/VBA_Challenge_2017.PNG)
![Speed of refactored macro for 2018](/Resources/VBA_Challenge_2018.PNG)

## Summary
### Pros and Cons of Refactoring Code
- Refactoring the original code created a macro that ran considerably faster while accomplishing the same analysis. 
- Refactoring the code took time and consideration of all the factors that could be tweaked to make it run more efficiently. 
- Editing functioning code always runs the risk of breaking it. 

### Pros and Cons of Original and Refactored Code
- The original code functioned as desired and only took less than a second to do so. 
- The refactored code ran even faster and was cleaner and easier to read. 
