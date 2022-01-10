# Stock Analysis with Excel

  

## **Overview of Project**
This project refactors the module 2 code, to loop through all data and collect the same information. Efforts were made to improve the efficiency of the refactored code, details of which are explained in the "Results" section below. 

### **Purpose**
The purpose of this project is to refactor the VBA code provided ("VBA_challenege.vbs") and to improve its executional efficiency.

## **Summary of Figures**
**a.** Analysis was completed on the source "Green Stocks" data using a reacfactored VBA script.

**b.** Analysis was completed and results confirmed accurate by matching the results from Module 2 exercises.

**c.** Results for original code, refactored code and optimized refactored code are shown in **_Figure 1_** to **_Figure 6_** 
  

![Figure 1](https://github.com/CR-HSDC/stock-analysis/blob/main/resources/VBA_Challenge_2017.png)

**_Figure 1_:** 2017 Optimized Refactored Code

![Figure 2](https://github.com/CR-HSDC/stock-analysis/blob/main/resources/VBA_Refactored_2017.png)

**_Figure 2_:** 2017 Unoptimized Refactored Code

![Figure 3](https://github.com/CR-HSDC/stock-analysis/blob/main/resources/VBA_Original_2017.png)

**_Figure 3_:** 2017 Original Code

![Figure 4](https://github.com/CR-HSDC/stock-analysis/blob/main/resources/VBA_Challenge_2018.png)

**_Figure 4_:** 2018 Optimized Refactored Code

![Figure 5](https://github.com/CR-HSDC/stock-analysis/blob/main/resources/VBA_Challenge_2018.png)

**_Figure 5_:** 2018 Unoptimized Refactored Code

![Figure 6](https://github.com/CR-HSDC/stock-analysis/blob/main/resources/VBA_Original_2018.png)

**_Figure 6_:** 2018 Original Code

![Figure 7](https://github.com/CR-HSDC/stock-analysis/blob/main/resources/CodeComparision.png)

**_Figure 7_:** Code Comparision

  

### **Results**
**a.** Optmiized Refactored Code for 2017 data took 0.539 Seconds to execute (*Figure 1*)

**b.** Original Code for 2017 data took 0.586 Seconds to execute (*Figure 2*)

**c.** Refactored Original Code for 2017 data took 0.621 Seconds to execute (*Figure 3*)

**d.** Optmiized Refactored Code for 2018 data took 0.543 Seconds to execute (*Figure 4*)

**e.** Original Code for 2018 data took 0.586 Seconds to execute (*Figure 5*)

**f.** Refactored Original Code for 2018 data took 0.617 Seconds to execute (*Figure 6*)

**g.** All data is reported to three (3) significant decimal places.

**h.** The Optimized Refactored code was the most efficient for both 2017 (*Figure 1*) and 2018 (*Figure 4*)

**g.** Code was optimized by elimination of a for loop in section *5a* of the *Run_OriginalCode* function. The SUM function was used within VBA to total the values between a starting and end row, eliminating the iterative computer cycles required by the FOR loop. This reduced computation time. The 		   code optimization can be seen from the red box highlights and comments in *Figure 7*




## **Summary**
-  **Advantages and Disadvantages of Refactoring Code**

	**a.** Advantages of refactoring code include time saving in writing and debugging original code.
	
	**b.** Disadvantages of refactoring code may include inefficiency of original code, particularly as it pertains to a specific use case, there may be redundant code that served a specific purpose in the original use case but is unncessary for the refactored use case.

	**c.** As the advantages and disdvantages detailed above, pertain to this particular use case, time was saved in developing the code for this application (advantage). The code had inefficiencies that needed to be optimized. As data sets grow, the computation time spent on these inefficiencies could also grow (disadvantage).  
