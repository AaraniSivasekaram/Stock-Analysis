# Stock Analysis with VBA

## Overview of Project

### The purpose of this analysis was to analyze different stock indexes per year, based on total daily volume and percentage return, to determine viable stocks to invest in.

## Results

- Overall stock performances between 2017 and 2018 differed significantly. In 2017 11 out of the 12 stocks achieved a positive return percentage by the end of the year, the one exception was TERP, which had a negative return percentage. Conversely, in 2018, 10 out of the 12 stocks, received a negative return percentage outcome by the end of the year. The two exceptions were ENPH and RUN, which both had positive return percentages in 2018. These findings suggest that 2018 was poor investment return year for this sample of stocks overall. 

- The execution times for the original script vs. the refactored script varied significantly. The original script ran the 2017 analysis in 0.727 seconds [https://github.com/AaraniSivasekaram/Stock-Analysis/blob/main/VBA_Module_2017.png] compared to the refactored script which ran the 2017 analysis in 0.141 seconds [https://github.com/AaraniSivasekaram/Stock-Analysis/blob/main/VBA_Challenge_2017.png]. The refactored script ran the analysis more than 5 times faster. The same results were seen in the 2018 analysis run times, the original script took 0.727 seconds of run time [https://github.com/AaraniSivasekaram/Stock-Analysis/blob/main/VBA_Module_2018.png] and the refactored script took 0.125 seconds of run time [https://github.com/AaraniSivasekaram/Stock-Analysis/blob/main/VBA_Challenge_2018.png]. 

## Summary 

1. Refactoring code is advantageous because it allows the code to become more efficient and therefore run faster. Another advantage of refactoring code is this step can improve the design of the code, to make the code easier to understand. A disadvantage of refactoring code is that it may introduce bugs or errors that the original script did not contain. As well, refactoring code is a time-consuming step, and this is a disadvantage. 

2. The disadvantage with refactoring the original VBA script was that I ran into errors when developing the refactored code. This step was time consuming and frustrating. Luckily, our wonderful TA @themarkfullton helped me sort out some kinks in my refactored code. The advantage of the refactored code in the VBA_Challenge.xlsm [https://github.com/AaraniSivasekaram/Stock-Analysis/blob/main/VBA%20Challenge.xlsm] is now the code can run with different sets of data. This new refactored code does not have many hard-coded/magic numbers and should be able to run analyses on different years for stocks values if these were added into the workbook.
