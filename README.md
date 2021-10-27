# stock-analysis with VBA
___

## Overview of Project

Steve is planning to do more research for his parents. He wants to expand the dataset to include the entire stock market.
The stock analysis VBA code we developed in module 2 might not work as well and take a long time to excecute for a larger amount of stocks.

Therefore the purpose of the project is to restructure the existing code in a way that improves the internal structure but doesn't change its external behavior.
We want to make the code faster and more efficient by taking fewer steps, using less memory, and improving the logic of the code.

We will measure with a timer in the code if refactoring indeed made the code run faster.
***

## Results

### Analysis

#### First step: 
I copied the starter code into VBA editor and reused the existing code to set up the Timer, create the user InputBox, activate the "All Stocks Analysis" worksheet, add the header, header row and initialize the array of tickers with the ticker values, activate the yearValue worksheet, and get the number of rows in the sheet to loop over.

  ![image](https://user-images.githubusercontent.com/91682586/138772963-ebb74438-ba95-406b-a745-85f50c30bb10.png)


#### 1a) I then created the tickerIndex that will later be used to access the arrays created in Step 1b) and set it to 0. 
    '1a) Create a ticker Index
    
    tickerIndex = 0

#### 1b) I created three output arrays for ticker Volumes, tickerStarterPrices, and tickerEndingPrices. The Datatype for ticker was set to Long, tickerStartingPrice and tickerEndingPrice were set to Single. 
    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
#### 2a) I created a for loop to initialize the tickerVolumes to zero.
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To RowCount 
    
        tickerVolumes(tickerIndex) = 0
               
    Next i
   
    

#### 2b) Created code that will loop over all the rows in the spreadsheet. 
            '2b) Loop over all the rows in the spreadsheet.
            
            For i = 2 To RowCount
    
  ##### 3a) This code increases the volume for the current ticker. 
            '3a) Increase volume for current ticker
            
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value   
            
  ##### 3b) This code checks if the current row is the first row with the selected tickerIndex and if true it assigns the current starting price to the tickerStartingPrice  variable.
            '3b) Check if the current row is the first row with the selected tickerIndex. If true assign starting price.
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
            'Assign Starting Price
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            End If
                     
   ##### 3c) This code checks if the current row is the last row with the selected ticker. If true it assigns the current ending Price to the tickerEndingPrice variable. If the next row’s ticker doesn’t match, the tickerIndex gets increased. 
   
            '3c) check if the current row is the last row with the selected ticker                        
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            'Assign Ending Price
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
   ##### 3d) This code increases the tickerIndex.
             'If the next row’s ticker doesn’t match, increase the tickerIndex.
            tickerIndex = tickerIndex + 1
                                    
            End If
            
  ####  The loop moves on to the next row.
        Next i    
   
       
   #### 4) This code loops through our arrays to output the Ticker, Total Daily Volume, and Return.
        'Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
            Cells(4 + i, 1).Value = tickers(i)
            Cells(4 + i, 2).Value = tickerVolumes(i)
            Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
            
        Next i 
  
  #### Formatting - I reused the existing Code to activate and format the Output worksheet, end the timer, and finish the Macro with End Sub. 
     'Formatting
        Worksheets("All Stocks Analysis").Activate
        Range("A3:C3").Font.FontStyle = "Bold"
        Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
        Range("B4:B15").NumberFormat = "#,##0"
        Range("C4:C15").NumberFormat = "0.0%"
         columns("B").AutoFit

         dataRowStart = 4
         dataRowEnd = 15

        For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
    'end timer    
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
    End Sub
  
### I finally run the stock analysis and confirmed that outputs for 2017 and 2018 were the same as they were in the module.

![AllStocksAnalysisResult2017](https://user-images.githubusercontent.com/91682586/138919189-bb3509e4-b0f1-4788-ad78-18974c309cc1.PNG)
![AllStocksAnalysisResult2018](https://user-images.githubusercontent.com/91682586/138919206-f06c5a41-65c3-416a-bd33-30df06e271ca.PNG)
   
### I also saved the run time of the new refactured code in the resources folder of this repository as VBA_Challenge_2017.png and VBA_Challenge_2018.png.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/91682586/138922812-5790652a-459d-4a80-92be-29a3dffd2909.PNG) 
![VBA_Challenge_2018](https://user-images.githubusercontent.com/91682586/138922823-b935ba35-0a9d-4225-9060-ea7b1be58a85.PNG)

### Summary
___
##### Comparison of the stock performance between 2017 and 2018 
***
2017 was a successful year for most stocks except TERP.  
2018 shows mostly negative returns except for ENPH and RUN which had again high returns.

##### Advantages and disadvantages of refactoring code
___

in General:
***
- Through refactoring code becomes easier to understand or read, faster, easier to to update and improve. This can save time and money in the future. 
- It helps the author coming back to read the code after a while as well as outside users.
- Refactoring can make the code more flexibel for other uses. 

- A disadvantage is that with complex code it might not be clear from the beginning how long exactly the completioin of the process might take and if there is a solution at all.   Because of the complexity you might end up spending a lot of time with little improvement in the end.

for the original VBA script:
***
- For our refactored stock-analysis code the main advantage is that it runs much faster. The new code has to loop though the data set only once instead of twelve times as in the original version.

  * For 2018 the original code needed almost a full second (0.8632813 seconds) to run versus an elapsed run time of 0.1367188 seconds for the refactored code. 
  * For 2017 the original code needed 0.8242188 seconds versus the refactored code taking only 0.109375 seconds.

Original Code:

![image](https://user-images.githubusercontent.com/91682586/139099800-3470444f-bb0a-4bdd-bb4e-ba4d3678cc4a.png)
![image](https://user-images.githubusercontent.com/91682586/139099440-09d61596-931d-40b3-935f-e4d42f939a0e.png) 


Refactored Code:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/91682586/139101037-81cbee9b-27f2-4bb6-bb98-f94eff86324e.PNG)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/91682586/139101041-684ed256-1b83-40c9-874e-1693f2d2bb51.PNG)



- A faster calculation is important for larger datasets.  

- The refactored code can be reused for other projects that require looping over items. 

- A disadvantage is that the new code is a bit more complex than the original code, e.g. requires understanding of arrays. 
