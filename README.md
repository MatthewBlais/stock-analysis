#STOCK ANALYSIS USING EXCEL VBA

## 1. Overview of Project: 
Steve's parents are looking to be Steve's first clients as he starts his new career as a financial manager. Steve's parents are eco conscious, so Steve put together a list of green stocks to be analyzed. After making code to analyze 12 green stock, the purpose of this challenge is to refactor, or edit, the code used so that we can analyze much more stocks for Steve.

## 2. Results

### Stock Performance
Looking at the two different years it is clear that 2017 was a better year for this group of 12 "Green" stocks than was 2018. in 2017, only 1/12 produced a negative return. In 2018 10/12 produced negative returns. See below for 2017 and 2018 results

https://github.com/MatthewBlais/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG

https://github.com/MatthewBlais/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG

### Execution Times

#### Before Refactoring
Before refactoring the data, the 2017 results ran in 0.5976563 seconds while 2018 took 0.5976563. 

https://github.com/MatthewBlais/stock-analysis/blob/main/Resources/VBA_Challenge_2017_before-refactoring.PNG

https://github.com/MatthewBlais/stock-analysis/blob/main/Resources/VBA_Challenge_2018_before-refactoring.PNG

#### After Refactoring
After refactoring the data, the 2017 results ran in 0.1015625 seconds while 2018 took 0.109375.

https://github.com/MatthewBlais/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG

https://github.com/MatthewBlais/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG

## 3. Summary

#### Advantages of refactoring code
Some advantages of refactoring code is that if done correctly it can become more efficient. This will execute similar goals in a faster time. This will also mean the code will take up less memory and run faster. The code is also going to be smaller and easier to read after a proper refactoring. 

#### Disadvantages of refactoring
If you have a working code, you will run the risk of ruining it while trying to refactor. That is why it is important to save work and use tools like Github so you can easily retrace your steps if a refactor goes poorly. You can also create a situation where you look for improvements but there are not any realistic things to be done. 

#### Original Vs. Refactored VBA Script
The Refactored VBA Script in this case did not return any different results, it just delivered them nearly 6 times faster than the original VBA. The code is also smaller and easier to follow on the refactored version. In reality, in this case the difference in small because the time to run are both under half a second. Although if we were to grow the list of stocks, the more added the better the difference in the codes will be.
