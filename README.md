![Illustration of the Stock Market](https://img.money.com/2022/05/News-Plunging-Stocks-401k.jpg?quality=60&w=1600 )
# Stock-Market-Analysis
## Description 
In this project, I'm analysing generated stock market data to uncover hidden trends and gain more insight into the dataset. Specifically, I'll create a script that loops through all the stocks for one year and outputs the following information:

- The ticker symbol
- Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
- The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
- The total stock volume of the stock.

Furthermore, I will also add functionality to my script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". 

## Usage
To see the VBA script, please navigate to "Hand in" folder and select "VBA script (challenge2)". It's recommended that you open this file via Excel VBA. 

## Installations
Please make sure you have VBA setup to run the code by following the visual instructions below:
### 1. Add the 'Developer' tab to the Excel ribbon
#### for Windows
Go to File Tab → Options → Customize Ribbon.\
In the main tab list, tick mark check box for the developer.
![Alt text](<Screenshot 2023-12-17 at 11.34.39 pm.png>)

Click 'OK'.

#### for Mac
Go to Excel Menu → Preferences.
![Alt text](<Screenshot 2023-12-17 at 11.36.26 pm.png>)

Click on Ribbon in "Sharing & Privacy Group" and then Click OK.
![Alt text](<Screenshot 2023-12-17 at 11.36.57 pm.png>)

Now, you will get a pop-up dialog box. In customization section, select Developer Tab & click OK.
![Alt text](<Screenshot 2023-12-17 at 11.37.24 pm.png>)

You should now see the 'Developer' tab in your ribbon. 
![Alt text](<Screenshot 2023-12-17 at 11.38.24 pm.png>)
### 2. Components of Visual Basic Editor
When you open the VBA editor, you should see the following display:
 ![Alt text](<Screenshot 2023-12-17 at 11.43.26 pm.png>)

## Stock market analysis
Our current dataset contains the following columns:
1. `ticker`
2. `date`
3. `open`
4. `high`
5. `low`
6. `close`
7. `vol`
![Alt text](<Screenshot 2023-12-17 at 11.53.29 pm.png>)

The following information was gathered from the current dataset using VBA:
1. `Ticker`: ticker names
2. `Yearly change`: yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
3. `Percent change`: the percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
4. `Total volume`: total stock volume of the stock.
5. `Greatest % Increase`: ticker which had the greatest percentage of increase.
6. `Greatest % Increase`: ticker which had the greatest percentage of decrease.
7. `Greatest Total Volume`: ticker which had the greatest total volume. 

![Alt text](<Screenshot 2023-12-17 at 11.29.25 pm.png>)

## Credits
Special thanks to the following individuals for their contributions:

-Jeffery Chieh-Liu (TA team): for helping me with conditional formatting 

-LiMei Hou (BCS tutor): for advicing me that I shouldn't have too many for-loops within a script because it would just slow down the execution process

-All other AskBCS helpers: for providing me with useful functions, such as the functions to convert numbers in scientific notation to normal numbers and converting decimals to percentages
