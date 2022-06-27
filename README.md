# VBA-Challenge

## Background
My client asked me to create a way to analyze thousands of stocks from multiple years using Microsoft Excel and VBA in order to determine the best investment options. 

## Overview of Project
The first macro I created analyzed the Total Daily Volume and Return for twelve potential stocks. 

![Original analysis](/Resources/Original_analysis.png) 

While this macro did what was needed, it was a bit slow. 

![Speed of original macro for 2017](/Resources/VBA_2017_original.PNG)
![Speed of original macro for 2018](/Resources/VBA_2018_original.PNG)

My client was interested in making the process more efficient so he could use it for future analyses. 

## Refactoring Results
In order to run the analysis more efficiently for my client, I refactored the original code so it iterated over all of the data for a given year once as opposed to multiple times.

![Refactored analysis](/Resources/Refactored_analysis.png)

This sped up the process considerably. 

![Speed of refactored macro for 2017](/Resources/VBA_Challenge_2017.PNG)
![Speed of refactored macro for 2018](/Resources/VBA_Challenge_2018.PNG)

## Summary

