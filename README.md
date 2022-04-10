# Stock_Analysis
Module 2 Stock_Analysis
Purpose:

The purpose of this assignment was to create a workbook that would analyze an entire data set quickly. While one was already made, it was not made as efficiently as it could have been. Here the challenge is to refactor the solution code, to loop through the entire data set. In order to compare whether the original was faster or the refactored one was.

Analysis of Data:

As stated above the data set has already been provided, with an already written code to process it. Here we are going to refactor the original code because of some limitations the original one had. The point of this is to make a fully functional code more efficient. To begin this, one had to assess what code was already made. Once that was figured out. A base refractory code was provided. But needed to be edited. Below it the process of filling in those blanks.

Step 1a:

    Dim tickerIndex As Single
    tickerIndex = 0

Step 1b:	

    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

Step 2a: 
          
    For i = 0 To 11
    ticker = tickers(tickerIndex)
    tickerVolumes(tickerIndex) = 0

        Next i
      
  Step 2b:
  
    For i = 2 To RowCount
    
Step 3a:

    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
  
      
Step 3b: 

    'If Then
     If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
     tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If
            
Step 3c:
        
     'If Then
      If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
      tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
     End If
            
 Step 3d: 
        
      if Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
	   tickerIndex = tickerIndex + 1
     End If
     
     Next i
    
    
Step 4: 
  
      For i = 0 To 11  
      Worksheets("All Stocks Analysis").Activate   
      Cells(4 + i, 1).Value = tickers(i)
      Cells(4 + i, 2).Value = tickerVolumes(i)
      Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

Results:

After completing the refractored analysis. For both 2017 and 2018. It did conclude that editing the data set was faster. For 2017 it did a complete code run in about 0.1679688 seconds. While the 2018 code ran 0.1640625 seconds see images blew.

<img width="281" alt="VBA_Challenge_17" src="https://user-images.githubusercontent.com/102453818/162643748-ffaa2d90-aaa2-47e3-91fa-9a500bf26c9b.png"> <img width="282" alt="VBA_challenge_18" src="https://user-images.githubusercontent.com/102453818/162643761-fcc4da51-17f0-4e4b-8b54-796462f5a905.png">


While refracting is beneficial, it does have its pros and cons. Some pros can include: Allowing for more flexibility and function. Another one might be maintainability. Which include the code itself being easier to read, as well as to maintain. More specifically, it’s cleaner and more organized. Now some of the disadvantages might be that it takes time. It’s definitely time consuming, and there is a greater chance for mistakes (Anarsolution, 2020). As well as possible loss of sanity. Especially when one gets stuck, it could take hours to figure out what went wrong. Which did happen on more than one occasion trying to refactor the challenge code.

Once it was running like it should have, there were definitely advantages to the refactored code. One being that it worked, and it was definitely faster. The main issue when running the original code, was it did not run for any amount of time. For example 2017 and 2018 ran for a total time of zero. As in nothing.  But once it was refactored,  2017 ran for 0.1679688 seconds, and 2018 ran for 0.1640625. Which could also be a disadvantage too. See while the original code seemed to be coded correctly, something in it was inherently wrong. Which was a disadvantage compared to the refracted one. However, one interesting thing did happen to the refactored code. When the code was run more than once, both 2017 and 2018 ran for the exact time 0.1640625 seconds. Until it was closed, and reopened which gave the current time frame.  But that being said, in this case, the only real downfall of a refracted code was that it took forever because it did not want to run. For some odd reason even correct code would make it bug out. Which meant more debugging was needed on correct code. Which was baffling. But what was even more interesting was restarting the program did make it work. So all in all,  there are definitely disadvantages and advantages to refactoring. But more importantly if the code doesn't run, and it is correct. Turn it off and then back on.
