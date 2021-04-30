# Stock-analysis using VBA

## Overview of the project

The purpose of this project was to help Steve to analyze a large amount of stocks data, I did this by creating a efficent code using VBA that can go through a dataset and extract for each stock, for any year, the stock daily volume and return percentage, while formating the cells so it can be easier to interpret.

## Results

### Initial Table

When I started this project, Steve had two dataset (2017 and 2018) that he wanted to analyze:

![Stocks_Initial_table](https://user-images.githubusercontent.com/81272629/116335944-530d0b80-a79d-11eb-831f-3eee05887fc8.png)

Each dataset contains above 3000 rows.

### Code

My code starts by asking the user to write the year he/she wants to analyze:

![Choosing_year](https://user-images.githubusercontent.com/81272629/116336299-ef371280-a79d-11eb-8f5d-031802ed3b26.png)

After the code has the year, it starts sorting through the data of that year (using arrays and loops) and extracting the Total Daily Volume and Return Percentage for each stock, then it format the Return column so the user can interpret the results easily. The resulting table if the user select the year 2018 is as follow: 

![Table_Results_2018](https://user-images.githubusercontent.com/81272629/116336658-8e5c0a00-a79e-11eb-89d7-594a6772f000.png)

### Original and Refactored Code 

For this project I made two codes that deliver the same results, the original one was a code that use nested loops to deliver the summary table, the code is as follow:

```
 Counting number of rows in 2018 sheet
 Worksheets(yearValue).Activate
 
 rowend = Cells(Rows.Count, 1).End(xlUp).Row
 
For i = 0 To 11

    Ticker = tickers(i)
    totalVolume = 0
    
Worksheets(yearValue).Activate
    
    For j = 2 To rowend

'Fin total volume of current ticker
    If Cells(j, 1) = Ticker Then
     totalVolume = totalVolume + Cells(j, 8).Value
     End If

'Find Starting price of current ticker
     If Cells(j, 1) = Ticker And Cells(j - 1, 1) <> Ticker Then
    Starting_price = Cells(j, 6).Value
    End If
    
'Find ending price for current ticker
     If Cells(j, 1) = Ticker And Cells(j + 1, 1) <> Ticker Then
    ending_price = Cells(j, 6).Value
    End If     
```
Although this code works properly, it has to loop 11 times around the whole dataset, which can be a problem if Steve want to use the same code to analyze 500 stocks instead of 11, thats why I made a code that use array and for loop to get the same result faster.

In the following pictures we can notice the difference between the code execution time of both codes for the year 2017.
 
Original code execution time:

![Old_Code_2017](https://user-images.githubusercontent.com/81272629/116634199-0a7e5b00-a921-11eb-9a31-dfd127273117.png)

Refactored code execution time:

![Refactored_Code_2017](https://user-images.githubusercontent.com/81272629/116634205-0d794b80-a921-11eb-8b90-302325335c7f.png)

## Summary 

### What are the advantages or disadvantages of refactoring code?

The biggest advantage of refactoring a code is that you probably will save time in the future if you want to use the code for a larger dataset. Sometimes you wil have to remake your code anyway because the one you have is not efficent enough for the task. Another advantage is that by doing it you learn in more depth how the programming language works.

One disadvantage of refactoring a code is the time you spend doing it, if the task that code perform is not very important, perhaps it is not worth it. Another disadvantage is that you can make your code too complex to other programmers to understand.

### How do these pros and cons apply to refactoring the original VBA script?

In my opinion, refactor this code was worth it, the result it is not a very complex script than can be understand by other programmers with ease. Also, it made the code efficient even if the data set is 50 times larger, and it made me a little bit better programmer by helping me understand different fuctions in VBA. 





