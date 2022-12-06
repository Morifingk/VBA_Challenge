#Morifing K.
# STOCK ANALYSIS WITH VBA + EXCEL 


## OVERVIEW: VBA Stock Analysis Project

### Purpose
In this project and analyisis, we’ll edit, or refactor, the Stock Market Dataset with VBA solution code to loop through all the data one time in order to collect an entire dataser. Then, we’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, we just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. 

### Analysis and Challenges
Here's a quick look at the Kickstarting Analysis and Challenges of this Project, including the following tasks:

- Prepare our dataser `VBA_Challenge.vbs` file for the project.
- Create our resources folder in **GitHub** to hold the run-time pop-up messages that we’ll screenshot after running refactored analyses for 2017 and 2018.
- Create and convert our `XLSM` file from `*.vbs` dataset that you used in this module as `VBA_Challenge.xlsm`.
- Add the VBA_Challenge.vbs script to the Microsoft Visual Basic editor.
- Use the steps **Refactor VBA code and measure performance** to add code where indicated by the numbered comments in the starter code file.

> Use your knowledge of VBA and the starter code provided in this Project to refactor the VBA Script dataset so we loop through the data one time and collect all of the information.

#### Our Challenge Data Background
> Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

> In this challenge, you’ll edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, you’ll present a written analysis that explains your findings.

> Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job.

## RESULTS: Refactor VBA Code and Measure Performance
 
### Deliverable Requirements, Code Examples, Compare Stock Performance and Timestamp procedure below:

**1. The `tickerIndex` is set equal to zero before looping over the rows.**

> Created a `tickerIndex` variable and set it equal to zero before iterating over all the rows. Will use this `tickerIndex` to access the correct index across the four different arrays on VBA Code: the tickers array and the three output arrays created on next requierement.


<img width="424" alt="Screen Shot 2022-12-06 at 3 38 06 AM" src="https://user-images.githubusercontent.com/104086409/205861608-5611bb3e-ad08-4df5-836e-a85ccc0d991f.png">


**2. Arrays are created for `tickers`, `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices`.**

> Created three output arrays: `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices`.
> In our VBA code, the `tickerVolumes` array should be a **Long** data type.
> But in our VBA code the `tickerStartingPrices` and `tickerEndingPrices` arrays should be a **Single** data type.


<img width="375" alt="Screen Shot 2022-12-06 at 3 39 04 AM" src="https://user-images.githubusercontent.com/104086409/205861815-3b5f9ebe-0dc2-4384-8cdd-61260dacadb4.png">


**3. The `tickerIndex` is used to access the stock ticker index for the `tickers`, `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices` arrays.**

> Created a for loop to initialize the `tickerVolumes` to **zero**. 
> And if the next row’s ticker doesn’t match, increase the `tickerIndex`.


<img width="415" alt="Screen Shot 2022-12-06 at 3 40 15 AM" src="https://user-images.githubusercontent.com/104086409/205862049-2c543415-84c3-4d5e-a83b-bcda5b1f9cc3.png">


**4. The script loops through stock data, reading and storing all of the following values from each row: `tickers`, `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices`.**

> Created a **loop** that will loop over all the rows in the spreadsheet.
> Inside the **loop**, we created a script that increases the current `tickerVolumes` **(stock ticker volume)** variable and adds the ticker volume for the current stock ticker.


<img width="602" alt="Screen Shot 2022-12-06 at 3 51 59 AM" src="https://user-images.githubusercontent.com/104086409/205864646-765f0f0b-39f1-4d34-a35f-f506b735dea1.png">


**Stored values from** `tickerStartingPrices` **and** `tickerEndingPrices`

> Created an **if-then** statement to check if the current row is the first row with the selected `tickerIndex`. If it is, then assign the current closing price to the `tickerStartingPrices` and `tickerEndingPrices` variable.


<img width="538" alt="Screen Shot 2022-12-06 at 3 53 01 AM" src="https://user-images.githubusercontent.com/104086409/205864900-04d8df7e-891a-493a-b7f2-88a3534da2d4.png">


**5. Code for formatting the cells in the spreadsheet is working.**

> We make positive returns green and negative returns red, to be a lot easier to determine which stocks did well and which ones didn't. Added some formatting based on the values of the returns. 


<img width="454" alt="Screen Shot 2022-12-06 at 3 56 25 AM" src="https://user-images.githubusercontent.com/104086409/205865719-f58af995-f240-4e78-9f69-1ad75927ae84.png">


**6. There are comments to explain the purpose of the code.**

> Adding **Comments** is requiered, as a **Best Practices for Writing Super Readable Code** such, 

- Commenting & Documentation, 
- Consistent Indentation, 
- Avoid Obvious Comments. 
- Code Grouping,
- Consistent Naming Scheme,
- DRY (Don't Repeat Yourself) Principle, 
- Avoid Deep Nesting,
- Limit Line Length, etc...



![name-of-you-image](https://github.com/emmanuelmartinezs/stock-analysis/blob/master/data_files/resources/Comments%20to%20explain%20the%20purpose%20of%20the%20code.PNG?raw=true)



**7. The outputs for the 2017 and 2018 stock analyses in the `VBA_Challenge.xlsm` workbook match the outputs from the AllStockAnalysis in the module**

> Finally, we run the stock analysis, to confirm that our stock analysis outputs for 2017 and 2018 are the same as dataset example provided (as shown in the images below, named **Dataset Examples Provided**). In adition, in our resources folder and below you can see the final Stock Analysis Results named, **Final VBA Analysis 2017 and 2018** save the pop-up messages showing elapsed run time for the refactored code as VBA_Challenge_2017.png and VBA_Challenge_2018.png. Then, save the changes to your workbook..

***Dataset Examples Provided***

<img width="631" alt="Screen Shot 2022-12-06 at 3 57 12 AM" src="https://user-images.githubusercontent.com/104086409/205865915-7c41c27a-552f-4c47-b590-1e26e630ef46.png">


> Below our Final VBA Analysis PNGs,


***Final VBA Analysis 2017***


<img width="229" alt="Screen Shot 2022-12-06 at 3 58 15 AM" src="https://user-images.githubusercontent.com/104086409/205866145-2a4c3071-a25e-4329-bdd1-bc9fca759fa3.png">


***Final VBA Analysis 2018***


<img width="229" alt="Screen Shot 2022-12-06 at 3 59 31 AM" src="https://user-images.githubusercontent.com/104086409/205866477-ee2c8373-4cb5-4247-92de-6b9d7d2e4832.png">


**8. The pop-up messages showing the elapsed run time for the script are saved as `VBA_Challenge_2017.png` and `VBA_Challenge_2018.png`**

> Running our fully 2017 and 2018 data stock analysis gave us an elapsed run time for each year, below our results.


***Time on VBA_Challenge_2017.PNG***


<img width="229" alt="Screen Shot 2022-12-06 at 3 58 15 AM" src="https://user-images.githubusercontent.com/104086409/205866145-2a4c3071-a25e-4329-bdd1-bc9fca759fa3.png">


***Time on VBA_Challenge_2018.PNG***

<img width="229" alt="Screen Shot 2022-12-06 at 3 59 31 AM" src="https://user-images.githubusercontent.com/104086409/205866477-ee2c8373-4cb5-4247-92de-6b9d7d2e4832.png">


## SUMMARY: Our Statement:

### Deliverable with detail analysis:
**1. What are the advantages or disadvantages of refactoring code?**

You need to perform code refactoring in small steps. Make tiny changes in your program, each of the small changes makes your code slightly better and leaves the application in a working state.

**Disadvantages:**

> - A long procedure may contain the same line of code in several locations, you can change the logic to eliminate the duplicate lines.
> - A logical structure may be duplicated in two or more procedures (possibly via copy & paste coding). When detected, this logic is best moved to a new function and called from the other functions.
> - A complex unstructured code is usually best to split in several functions. 
> - Refactoring process can affect the testing outcomes. 


**Advantages:**
> - Logical errors easily appear in well structure code that contains nested conditionals and loops. 
> - In our case, using Excel flow displays program logic in a more comprehensible manner, not tied to the order that the underlying code is written.
> - VBA interpretation (Excel) of code can reveal patterns that are not easy to see in the source.

**2. How do these pros and cons apply to refactoring the original VBA script?**

> Improving or updating the code without changing the software’s functionality or external behavior of the application is known as code refactoring.
Now, let's think about something, **What happens after a couple of days or months yo need to troubleshoot your code? Is it complicated? Is it hard to understand?** If yes then definitely you didn’t pay attention to improve your code or to restructure your code. 

***We need to consider the code refactoring process as cleaning up the orderly house.*** 
*Unnecessary clutter in a home can create a chaotic and stressful environment.* - The same goes for written code. 

A clean and well-organized code is always easy to change, easy to understand, and easy to maintain. You can avoid facing difficulty later if you pay attention to the code refactoring process earlier.
