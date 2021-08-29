
# <div align="center">__Stock Analysis__</div>

## <div align="center">__Overview of Project__</div>
The main goal of this macro-enabled workbook is to provide an analysis tool for any previous or future stocks. This tool aims to provide a quick overview of any stocks' total volume traded within a given period and its percent returns. 

We have built two versions of this project, both provide said analysis with the second version being more efficient. Our goal is to provide insight on why _Refactoring_ code is important in analyzing data.

## <div align="center">__Results__</div>

<div align="center"> 

![alt-text-1](https://raw.githubusercontent.com/RobC30/stock-analysis/main/Resources/stock2017.png) ![alt-text-2](https://raw.githubusercontent.com/RobC30/stock-analysis/main/Resources/stock2018.png) 
 <br>
 ___2017 (left) vs. 2018 (right)___
</div>

For the stock performance analysis, 2017 was generally better vs. 2018. In 2017, outstanding performances can be seen with 'DQ' & 'SEDG' nearly gaining a 200% increase in return and almost every other stock saw positive increases except for 'TERP.' In terms of volume traded, top performer is 'SPWR' with 'FSLR' running in second, which may mean these were popular with the public. With Steve's parents pestering him about 'DQ's stock, looking at 2017 alone, it might be a good idea to invest in it however a year's worth of data might not be enough.

In 2018, almost every stock had a negative outlook except for 'ENPH' & 'RUN.' Every stock either saw an increase or decrease in traded volumes but these two said stocks had the largest increases which could be attributed to its success. If these stocks had positive outlooks for 2 consecutive years, we can definitely suggest to invest in 'ENPH' & 'RUN' to Steve's parents.

## Original vs. Refactored
The key difference between the two VBA scripts is the use of indices and arrays. On the first code, we loop through each tickers on 1 (one) array to perform the analyis. The code runs smoothly giving accurate results in analyzing returns & total volumes traded on any given stock. 

## Original Macro
It was given a timer on how quickly it can run _3,000_ lines of data. For the year 2017, it ran in __0.6445313 seconds__ and for the year 2018, it ran in __0.7128906__ seconds. Now this might seem a good time in running the analysis, however if given a bigger set of data to handle it might become inefficient.

_See image below for original code details:_
<br>
![image](https://raw.githubusercontent.com/RobC30/stock-analysis/main/Resources/code.PNG)
<br>

![alt-text-1](https://raw.githubusercontent.com/RobC30/stock-analysis/main/Resources/orig_runtime2017.png) ![alt-text-2](https://raw.githubusercontent.com/RobC30/stock-analysis/main/Resources/orig_runtime2018.png)

<br>

## Refactored Macro
For the refactored VBA script, we still went through each tickers but added new arrays for the code by using nested loops & assigning variables to said arrays. We introduced three new arrays for _Volumes Traded, Staring & Ending Prices_. For the years 2017 & 2018, it ran in __0.0859375__ seconds. This runs the analyis in a fraction of the original time establishing a more efficient code. It will analyze bigger data sets on a shorter time.

<br>

_See image below for refactored code:_

![image](https://raw.githubusercontent.com/RobC30/stock-analysis/main/Resources/code2.PNG)

<BR>

_See images below for refactored code runtimes:_

![alt-text-1](https://raw.githubusercontent.com/RobC30/stock-analysis/main/Resources/refactored_runtime2017.png) ![alt-text-2](https://raw.githubusercontent.com/RobC30/stock-analysis/main/Resources/refactored_runtime2018.png)


## <div align="center">__Summary__</div>
- Refactoring code gives the possibilty of running it more efficient, giving shorter runtimes on analyzing data. Refactoring will definitely help on larger sets of data but dealing with a small dataset, it might not be worth a coder's time. The given dataset is a great example, with the original code providing results in less than a second even if the refactored code slashes the runtime by fractions.
- Refactoring an existing _VBA script_ can give you a side-by-side comparison between it & the existing script. One can create a roadmap on which areas of code could be improved increasing efficiency by creating more arrays, cleaner coding & etc. However, if one does not fully understand the syntax of an existing code, there will be a chance of making a perfectly running script into an unstable & unusable one.

