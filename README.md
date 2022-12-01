# VBA_Challenge #Morifing Koné
OVERVIEW: VBA Stock Analysis Project
Purpose
In this project and analyisis, i’ll edit, or refactor, the Stock Market Dataset with VBA solution code to loop through all the data one time in order to collect an entire dataser. Then, i'll determine whether refactoring your code successfully made the VBA script run faster. Finally, I just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.

Analysis and Challenges
Here's a quick look at the Kickstarting Analysis and Challenges of this Project, including the following tasks:

Prepare the dataser VBA_Challenge.vbs file for the project.
Create the resources folder in GitHub to hold the run-time pop-up messages that we’ll screenshot after running refactored analyses for 2017 and 2018.
Create and convert our XLSM file from *.vbs dataset that you used in this module as VBA_Challenge.xlsm.
Add the VBA_Challenge.vbs script to the Microsoft Visual Basic editor.
Use the steps Refactor VBA code and measure performance to add code where indicated by the numbered comments in the starter code file.
Use your knowledge of VBA and the starter code provided in this Project to refactor the VBA Script dataset so we loop through the data one time and collect all of the information.

Our Challenge Data Background
Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

In this challenge, i'll edit, or refactor, the solution code to loop through all the data one time in order to collect the same information that you did in this challenge. Then, i’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, i’ll present a written analysis that explains my findings.

Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job.

RESULTS: Refactor VBA Code and Measure Performance
Deliverable Requirements, Code Examples, Compare Stock Performance and Timestamp procedure below:
1. The tickerIndex is set equal to zero before looping over the rows.

Created a tickerIndex variable and set it equal to zero before iterating over all the rows. Will use this tickerIndex to access the correct index across the four different arrays on VBA Code: the tickers array and the three output arrays created on next requierement.

name-of-you-image

2. Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

Created three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices. In our VBA code, the tickerVolumes array should be a Long data type. But in our VBA code the tickerStartingPrices and tickerEndingPrices arrays should be a Single data type.

name-of-you-image

3. The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays.

Created a for loop to initialize the tickerVolumes to zero. And if the next row’s ticker doesn’t match, increase the tickerIndex.

name-of-you-image

4. The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

Created a loop that will loop over all the rows in the spreadsheet. Inside the loop, we created a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker.

name-of-you-image

Stored values from tickerStartingPrices and tickerEndingPrices

Created an if-then statement to check if the current row is the first row with the selected tickerIndex. If it is, then assign the current closing price to the tickerStartingPrices and tickerEndingPrices variable.

name-of-you-image

5. Code for formatting the cells in the spreadsheet is working.

We make positive returns green and negative returns red, to be a lot easier to determine which stocks did well and which ones didn't. Added some formatting based on the values of the returns.

name-of-you-image

6. There are comments to explain the purpose of the code.

Adding Comments is requiered, as a Best Practices for Writing Super Readable Code such,

Commenting & Documentation,
Consistent Indentation,
Avoid Obvious Comments.
Code Grouping,
Consistent Naming Scheme,
DRY (Don't Repeat Yourself) Principle,
Avoid Deep Nesting,
Limit Line Length, etc...
name-of-you-image

7. The outputs for the 2017 and 2018 stock analyses in the VBA_Challenge.xlsm workbook match the outputs from the AllStockAnalysis in the module

Finally, we run the stock analysis, to confirm that our stock analysis outputs for 2017 and 2018 are the same as dataset example provided (as shown in the images below, named Dataset Examples Provided). In adition, in our resources folder and below you can see the final Stock Analysis Results named, Final VBA Analysis 2017 and 2018 save the pop-up messages showing elapsed run time for the refactored code as VBA_Challenge_2017.png and VBA_Challenge_2018.png. Then, save the changes to your workbook..

Dataset Examples Provided

name-of-you-image

Below our Final VBA Analysis PNGs,

Final VBA Analysis 2017

name-of-you-image

Final VBA Analysis 2018

name-of-you-image

8. The pop-up messages showing the elapsed run time for the script are saved as VBA_Challenge_2017.png and VBA_Challenge_2018.png

Running our fully 2017 and 2018 data stock analysis gave us an elapsed run time for each year, below our results.

Time on VBA_Challenge_2017.PNG

name-of-you-image

Time on VBA_Challenge_2018.PNG

name-of-you-image

SUMMARY: Our Statement:
Deliverable with detail analysis:
1. What are the advantages or disadvantages of refactoring code?

You need to perform code refactoring in small steps. Make tiny changes in your program, each of the small changes makes your code slightly better and leaves the application in a working state.

Disadvantages:

A long procedure may contain the same line of code in several locations, you can change the logic to eliminate the duplicate lines.
A logical structure may be duplicated in two or more procedures (possibly via copy & paste coding). When detected, this logic is best moved to a new function and called from the other functions.
A complex unstructured code is usually best to split in several functions.
Refactoring process can affect the testing outcomes.
Advantages:

Logical errors easily appear in well structure code that contains nested conditionals and loops.
In our case, using Excel flow displays program logic in a more comprehensible manner, not tied to the order that the underlying code is written.
VBA interpretation (Excel) of code can reveal patterns that are not easy to see in the source.
2. How do these pros and cons apply to refactoring the original VBA script?

Improving or updating the code without changing the software’s functionality or external behavior of the application is known as code refactoring. Now, let's think about something, What happens after a couple of days or months yo need to troubleshoot your code? Is it complicated? Is it hard to understand? If yes then definitely you didn’t pay attention to improve your code or to restructure your code.

We need to consider the code refactoring process as cleaning up the orderly house. Unnecessary clutter in a home can create a chaotic and stressful environment. - The same goes for written code.

A clean and well-organized code is always easy to change, easy to understand, and easy to maintain. You can avoid facing difficulty later if you pay attention to the code refactoring process earlier.
