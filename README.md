# StockDataAnalysis

 The data for this project consists of information from multiple stocks from 2018 to 2020. The data contains each stock's name, its open price, its close price, the high price, the low price, and the volume.


 The Excel VBA script I have created creates a summary table for each year. This summary table consists of four columns. The first column (Ticker) gives the name of each unique stock name that appears in the respective year's sheet. The second column (Yearly Change) calculates the difference between the close price and the open price for each individual stock name. The third column (Percent Change) calculates the percentage of difference between the opening price for each stock divided by the closing price in a given year. My script also highlights values in the 'Yearly Change' and 'Percent Change' columns green or red. Green indicates a positive change, and red indicates a negative change. This makes it easier to read whether each stock has increased or decreased in value over the course of a particular year. The final column in this summary table shows the cumulative stock volume throughout the year for each stock.


 My VBA script also locates each year's greatest percent increase and the stock's name associated with the highest value. Likewise, it also identifies the greatest percent decrease and provides the stock's name for this as well. This secondary summary table also includes the sgreatest aggregated stock volume for stock.
