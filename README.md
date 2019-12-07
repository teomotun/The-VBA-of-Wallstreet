# VBA Exercise - The VBA of Wall Street

	# Due to size of the source data, the XML file has been stored on Google drive, https://drive.google.com/file/d/0ByPvWDiJNfm7el81aXJiUkM1LTg/view

	# 3 levels of VBA scripting have been created to ease “Full” takes roughly 1.5 minutes to run while “Lite” will have the shortest run time.

* Script buttons are located on each sheet and will require Macros to be activated within Excel.
* Multiple_year_stock_data Tuttle.xml is a very large data file and will take time to load. A separate word file has been included to view solution code. 

### Objective

* The script will loop through all the stocks and take the following info:

	* Yearly change from the stock price at the beginning of the year to the price at the end of the year.

	* The percent change in that stock’s price from the beginning of the year to the end.

	* The total volume of the stock traded over the year with ticker symbol

* Apply conditional formatting to highlight a yearly increase in the stock price in green, and a decrease in the stock price in red.

* Identify the stock with the greatest percentage increase in its price; the stock that suffered the greatest percentage decrease in its price; and finally, identify the stock with the greatest trade volume.

* The script should cover every worksheet by running it once
