# Utilizing VBA to analyze Green Stocks
## Project Overview
This project is to assist Steve and his parents to support DAQO, a new energy company. They are very supportive of Green energy and Steve wanted to analyze other Green energy stocks, in addition to DAQO
### Purpose
To develope and analyze data on Microsoft Excel with VBA. Data set includes many different stocks for year 2017 to year 2018, using basic functions such as 
- Loops
- Pop-Ups
- Inputs and MsgBox
- Cell Formating
- Code Refactoring

## Results
After analyzing the dataset for DAQO from year 2017 and year 2018, it appears that this stock has gone down in 2018, by almost 60%.
![image](https://user-images.githubusercontent.com/104419959/187312430-d9eae765-4e45-4e5a-9642-cee6ae7138b7.png)
#### Run time
The initial run on year 2018 took about 1.16 seconds

![VBA_Challenge_2018_Original](https://user-images.githubusercontent.com/104419959/187312700-7cddf5f0-c82c-479e-966a-4c7517e12e77.png)

After refactoring, the speed of the run increased by almost 5 times, aobut 0.2 seocnds.

![VBA_Challenge_2018_Refactored](https://user-images.githubusercontent.com/104419959/187312880-611b3f3d-cde8-480d-9b69-6ac3d6f90c6d.png)





The main reason for this increase of speed is how we loop the data. 
From the old code, we used nested loops, which took time to process the inner loop and the outer loop, and then output the result. 
![image](https://user-images.githubusercontent.com/104419959/187313225-542912ed-7122-42c9-9f6f-a2e5392efa19.png)
After adjustments, we only used one loop to process the data. 
![image](https://user-images.githubusercontent.com/104419959/187313328-b06a53fa-bc1a-4156-b412-a9c779afca44.png)

## Summary
### Pros and Cons of refactoring codes
The biggest and most obvious advantage of refactoring codes is that it allows us to save time when processing the data. We are able to edit and modify our codes to look neater, and also edit any mistakes that were found. It allows us to make sense of the code and we can use it as an entry point to working with existing code.  
The only disadvangtage that I am seeing is that it takes up time to recode and ensuring that it will function properly. 

The refactored VBA codes are much more efficient and requires lesser memory to store. The new codes functions perfectly, even better than the original codes where we had to use nested loops. There are no disadvantage to the refactored codes from what I am seeing. 
