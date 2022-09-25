# VBA-challenge

Background:
You are well on your way to becoming a programmer and Excel expert! In this homework assignment, you will use VBA scripting to analyze generated stock market data. Depending on your comfort level with VBA, you may choose to challenge yourself with a few of the bonus Challenge tasks.

Challenge:
Create a script that loops through all the stocks for one year and outputs the following information:
1.	The ticker symbol.
2.	Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
3.	The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
4.	The total stock volume of the stock.

Other points to consider includes
1.	Make sure to use conditional formatting that will highlight positive change in green and negative change in red.

Solution:
Steps taken to solve the challenge:
1.	Declare Worksheet and other variables
2.	Loop through all the worksheets
3.	Count the number of number of rows in the first column
4.	Set header for the values
5.	Set counter and open_value
6.	Loop through the rows by the ticker names
7.	Search for when the value of the next cell is different from the current cell
8.	Increase the counter when the condition meets and passing the corresponding close_value and ticker name
9.	Pass the ticker names to the excel
10.	Calculate the yearly change and passing the value to excel
11.	Calculate the percent change and passing them to the excel
12.	Calculate stock volume
13.	Set color code for the yearly change
14.	Clear the values
15.	Calculate the stock volume
16.	Fetch the open_value
17.	Label the Summery table headers
18.	Find the last row of the table
19.	Loop through the summary table
20.	Find the Max percent change and the corresponding ticker name
21.	Find the Min percent change and the corresponding ticker name
22.	Find the Max Volume and its ticker name

Bonus Assignment:
Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". 

Solution:
Loop through each worksheet to find the greatest increase and decrease as well greatest total volume.

Reference Links:
https://stackoverflow.com/
