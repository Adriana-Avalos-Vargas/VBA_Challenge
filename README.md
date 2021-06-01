# VBA_Challenge - The VBA of Wall Street (Stock market analysis)

## Introduction

According to Investopedia *"The stock market refers to the collection of markets and exchanges where regular activities of buying, selling, and issuance of shares of publicly-held companies take place. Such financial activities are conducted through institutionalized formal exchanges or over-the-counter (OTC) marketplaces which operate under a defined set of regulations."*  The stock markets are essentials components of a free-market economy because they enable democratized access to trading and exchange of capital for investors of all kinds. Therefore, in order to understand how does the actions of an enterprise or company behave and therefore to profit in the financial markets.

The idea is to design a VBA script to perform an analysis of real stock market data by using test data and then real stock data.

### Stock market analysis

![stock Market](Images/stockmarket.jpg)

By using a sample of the stock market data within the file **`alphabetical_testing.xlsx`** a VBA code is developed. This is done to ensure that the data set is small and therefore it allows a faster test (around 3 to 5 minutes). 

The script loops through all the stocks for one year and output the following information.

* The ticker symbol.

* Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

* The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

* The total stock volume of the stock.

* It also includes conditional formatting that will highlight positive change in green and negative change in red.

* The result is similar to the one in the image below.

![moderate_solution](Images/moderate_solution.png)

### Additional Information

1. Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume". The solution will look as follows:

![hard_solution](Images/hard_solution.png)

2. Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.

### Other Considerations

* Use the sheet `alphabetical_testing.xlsx` while developing your code. This data set is smaller and will allow you to test faster. Your code should run on this file in less than 3-5 minutes.

* Make sure that the script acts the same on each sheet. The joy of VBA is to take the tediousness out of repetitive task and run over and over again with a click of the button.

## Submission

* To submit please upload the following to Github:

  * A screen shot for each year of your results on the Multi Year Stock Data.

  * VBA Scripts as separate files.

* After everything has been saved, create a sharable link and submit that to <https://bootcampspot-v2.com/>.

- - -

### Copyright

Trilogy Education Services © 2019. All Rights Reserved.

