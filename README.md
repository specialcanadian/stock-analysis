# VBA Stock Analysis

## Overview of Project

Steve requested A VBA program he could use to study DQ’s stock performance during 2017 and 2018. After using the program to analyze DQ stock, Steve wanted to do a similar analysis on the rest of the stock market. To handle the much larger amount of data, I refactored the code with the goal of having it run more efficiently.  

##Results

### Code Analysis
The code was refactored following the specific Canvas instructions and runs much faster than before. Here are the current time results for the refactored code, and the original (Fig.1, Fig.2).

 
![alt text](https://github.com/specialcanadian/stock-analysis/blob/main/Resources/GithubImgRefactoredCode.png?raw=true)

Fig.1: Refactored code runtime for 2017 and 2018


![alt text](https://github.com/specialcanadian/stock-analysis/blob/main/Resources/GitHubImgOldCode.png?raw=true)  

Fig.2: Original code runtime for 2017 and 2018

As we can see, the refactored code is about 0.4s faster than the original.

The refactored code’s superior performance comes from removing a nested for loop.   



## Summary

###


