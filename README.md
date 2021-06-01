# VBA_Challenge - The VBA of Wall Street (Stock market analysis)

## Introduction

According to Investopedia *"The stock market refers to the collection of markets and exchanges where regular activities of buying, selling, and issuance of shares of publicly-held companies take place. Such financial activities are conducted through institutionalized formal exchanges or over-the-counter (OTC) marketplaces which operate under a defined set of regulations."*  The stock markets are essentials components of a free-market economy because they enable democratized access to trading and exchange of capital for investors of all kinds. Therefore, in order to understand how does the actions of an enterprise or company behave and therefore to profit in the financial markets.

The idea is to design a VBA script to perform an analysis of real stock market data by using test data and then real stock data.

### Stock market analysis

![stock Market](Images/stockmarket.jpg)

By using a sample of the stock market data within the file **`alphabetical_testing.xlsx`** a VBA code is developed. This is done to ensure that the data set is small and therefore it allows a faster test (around 3 to 5 minutes)

The script loops through all the stocks for one year and output the following information.

* The ticker symbol.

* Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

* The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

* The total stock volume of the stock.

* It also includes conditional formatting that will highlight positive change in green and negative change in red.

* The result is similar to the one in the image below.

![moderate_solution](Images/moderate_solution.png)

### Additional Information

Since there are other important indicators when studying the strength of stocks in the market and the VBA code must return the following data:

1. The stock with the "Greatest % increase",

2. The stock with  "Greatest % decrease" 

3. The stock with "Greatest total volume". 

Such indicators are presentes as shown below.

![hard_solution](Images/hard_solution.png)

**Note: Once the VBA scripts works on the short versión data it is important to guarantee thar the VBA script will  run on every worksheet of the complete data just by running the VBA script once. This is important since each worksheet represents a year. 

### Results

The VBA script is uploaded into the github repository but not the databases. The results are presented in a PDF file named *"VBA _challenge_screen_shots_AAV"*.
### Copyright

Trilogy Education Services © 2019. All Rights Reserved.
