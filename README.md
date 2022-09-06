# Overview of Project

## Background

The original VBA script analyzing the dataset of dozen stocks in 2017 and 2018 worked well. However, it might not work well or might take a long time to execute if the dataset includes the entire stock market over the last few years.   

## Purpose

This project is to refactor the original code to loop through all the data one time in order to collect the same information and then determine whether refactoring the code successfully made the VBA script run faster.

# Results

## Stock performance

  * The refactored script and original script analyzed the same dataset and provided the same stock performance in 2018. For example, for stock "AY" in 2018, both scripts calculated the same values for The Total Daily Volume and Return, which are 83,079,900.00 and -7.3%
  
  * Below are images of stock performance in 2017 and 2018 analyzed by the original VBA script and the refactored script:
  
    ![stock performance_2017](https://user-images.githubusercontent.com/110554264/188664834-8ac00f14-1286-4d4d-b823-626eac5350c1.png)     
    ![stock performance_2018](https://user-images.githubusercontent.com/110554264/188664943-b908e91d-5c55-494a-a9de-6762282f25f3.png)
    
## Execution times

  * The refactored script takes less time to execute than the original script. Particularly, to analyze the stock performance in 2018, the original script takes 1.58 seconds while the refactored script takes about 0.14 seconds.
  
  * The refactored script uses arrays to store values of Volumes, Starting Price, and Ending Price of stocks instead of using single variables in the original script. This change helps the script run faster than the original script.  
  
    ![arrays](https://user-images.githubusercontent.com/110554264/188676144-607af0b2-b2f1-4715-bda6-1b86a98f56e7.png)

  * Below are images showing elapsed run time of the original script and refactored script to analyze the stock performance in 2018:
  
    Elapsed run time of the original script:         
    ![elapsed run time_2018_original](https://user-images.githubusercontent.com/110554264/188670274-7a669c01-f25e-4bbe-96d9-e1c2b945a945.png)
    
    Elapsed run time of the refactored script:    
    ![VBA_Challenge_2018](https://user-images.githubusercontent.com/110554264/188670501-992f5154-0900-4aab-a47d-442759290503.png)
    
# Summary

In general, refactoring has both advantage and disadvantages. Refactoring is the way to improve the code quality. The code after refactoring is fresher, easier to understand or read, and easier to maintain and enhance. Refactoring also helps in executing the program faster. However, refactorting consumes much more time not only for refactoring process but also for testing the new solution to make sure that the new solution covers all business scenarios as the original code does. If it is not tested properly, refactored code may introduce new defects.

For this project, refactored code successfully made the VBA script run faster. The code is restructed and optimized without changing its behavior. If the process went wrong and we did not reach the goal of improving the execution time, it would take more time for us to work on the refactoring process.

    
