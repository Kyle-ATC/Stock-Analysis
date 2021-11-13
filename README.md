***GREEN STOCK ANALYSIS***
   ***Overview of Project***
       My good friend Steve has asked for my help in assisting his parent’s request for reviewing stocks, specifically green energy.  Steve has recently become a stockbroker and wants to help his parents by utilizing my skills in VBA (I told him this was a mistake, but he believes in me nonetheless) with his knowledge of the stock market.
What Steve and I are looking to do is build a code that will help in both the short term and long term.  The short term is to have a code to review the past two years of a few green energy stocks to see how “DQ” is performing, 
is it making his parents money
Are there other successful companies from the past two years
The long term is to have a code Steve can use to analyze many large quantities of stock in a reasonable amount of time during his busy workday.
   *Results*
With a refactored code, I would assume this process would take less than the approximate 1 second it took to run on my computer for each year.  I ran this three times, prior to refactoring to establish a good baseline. The original code had no real difference.  See below.
 - 2017
   - First run![This is an Image](https://drive.google.com/file/d/1tiKL8tAa0bibGswYVHzYQiu6ceWz7qeS/view?usp=sharing)
     -  1.03 seconds
   - Second run ![This is an Image](https://drive.google.com/file/d/1QM8euq1y_bIay5sMaNCY4_IY-qIKCUhm/view?usp=sharing)
     - 1.07 seconds
   - Third run ![This is an Image](https://drive.google.com/file/d/1eMT-g5JysesI98xLsvqY-e3pc8wQz42z/view?usp=sharing)
     - 1.00 seconds
 - 2018
   - First run ![This is an Image](https://drive.google.com/file/d/1690-BW5BrK-HoELaCzeT8GusBq73coS1/view?usp=sharing)
     -  1.27 seconds
   - Second run ![This is an Image](https://drive.google.com/file/d/1gKs9Vhn4m3NhL89buPGTjrZsSH8CupGg/view?usp=sharing)
     - 1.38 seconds
   - Thrid run ![This is an Image](https://drive.google.com/file/d/1eOs2p3DBCHQotrHLZmc4QiUM2oKN9y1j/view?usp=sharing)
     - 1.40 seconds
        
    *Solved*    
        Unfortunately, I have yet to make my code work.  I will likely double back on this project at some point and find why I am getting an error for the section 3b of the VBA code 
        ```
        If Cells(i - 1, 1).Value <> "tickers(tickerIndex)" And Cells(i, 1).Value = "tickers(tickerIndex)" Then
        ```
   - “refactored code images to be added later”


        What I was attempting to do, was to establish the arrays ahead of time in the code, so when the computations ran, it would know what it need to find 12 variables and then could plug each one into the equations for the next set of variables and formulas as they came, opposed to looping through all of the lines every time for each item in the array. The code refactoring revolved array naming the arrays earlier in the code, instead of identifying during the loops.
```
       Dim tickerVolumes(12) As Long
       Dim tickerStartingPrices(12) As Single
       Dim tickerEndingPrices(12) As Single”
```
    ***Summary***
       *Disadvantages of refactoring code*
          This particular code, like most codes, used nested loops.  This format of coding, while clear and established, takes a longer time to review all rows, columns, pages for each item in the array.  This adds a large amount of computational memory as well as can take a large amount of time to accomplish (depending on the size of the data file).  Other cons are what happened with me.  Trying to refactor the code may lead to introducing bugs, potentially losing data or parts of the equations as you change the structure of your code.  This can lead to errors and the code ultimately not running.  Another con to mention is, if you have a working code, refactoring will not produce a more correct answer.
       *Advantages of refactoring code*
          By refactoring and finding simpler, more sorted information ahead of time, you can increase the processing speed saving you time.  This more organized code will hopefully be easier to read and see the steps of what you are trying to achieve.  Additionally, this faster code will allow you to achieve results and increase productivity.

        *What happened to my refactor*
          My attempt on the refactored code to achieve a faster time is not really required.  With it being a small data set, and the total time elapsed is 1 second to run the code, there was not much of a change for this data set.  
          The cons unfortunately really had the upper hand.  I, unfortunately, was often stuck for long periods attempting to refactor the code that I would sometimes forget what I was trying to modify.  Luckily I was making periodic saves on my desk for version control that I was able to revert back.  I would also look at the looks I was making to see what I was trying.  I often moved no longer needed code, or code that was not working to the bottom of the page, after end sub to track these items when I would compare against previous versions.
