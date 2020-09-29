# ExcelAutomation
As we all end up using Excel a lot for any type of data analysis/preprocessing, it would be great to automate any Excel activity using Python and very powerful library Xlwings !!

# Steps to follow for using the Xlwings enabled sheet
1. Download the xlsm sheet and python (my_math) script and save in your local directory
2. Open the Excel and right click on button 'Calculate Numbers', select 'Assign macro', select 'math_calculate'** and save
3. Enter any number in A3 and ODD number is B3 (preferable less than 20)
4. Click on the button 

** Read the medium article to figure out how a Frozen Python (.exe) can be used instead of Python script (.py)

# Algorithm
1. There is an interesting property of numbers which when applied in a matrix (odd no * odd no) formation, it results in summation of numbers row-wise/column-wise to same number
2. After multiple iterations and experimentation, I was able to figure out the formula and hence generalized this beatiful property of numbers in Python. As part of reverse engineering -  for any given number and odd matrix table the numbers are arranged in such a fashion that it sums up to the same given number
3. This is coded in Python and integrated with Excel for a basic demo of Xlwings. Do try out and see the Math magic !!
