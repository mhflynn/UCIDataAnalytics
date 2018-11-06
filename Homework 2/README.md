# Unit 2 | Assignment - The VBA of Wall Street

* Please see the assignment README <https://www.github.com>

### Instructions 

* The file StackAnalysis.bas is an exported module from Execl. To test the program, import the file into the workbook file with stock data and run as follows :
   * Open the Multiple_year_stock_data.xlsx file
   * From the Developer tab, open the Visual Basic window
   * Ctl-Select "This Workbook" and import file  
   * Run the macro "StockAnalysis"
   * Check the results

### Description

 * For each worksheet in a file, this function will process a table
 containing stock price and volume data by ticker symbol. Each 
 worksheet is assumed to contain data for a single year. 
 A summary table will be created on the same spreadsheet, to the right
 of the original table, containing ticker symbol, gain($), gain(%) and total volume.
 Additionally a post processing step will be run on the summary table to :
   - Conditionally color cells with percent change, green for positive
     gains, red for negative gains for the year
   - Output an additional table with values for least gain, highest gain and highest 
     volume by ticker symbol

 * **Assumptions for input table :**
   - The first cell of the input table is "A1"
   - The input spans 7 columns "A" - "G"
   - The table spans a variable number of rows
   - The end of the table is the first row with empty cell in column "A"
   - The first row in the table contains headers as below
   
       |\<ticker\>| \<date\> | \<open\> | \<high\> | \<low\> | \<close\> | \<volume\> |
       |---|---|---|---|---|---|---|
       
   - The content of the header row may be checked, but it is assumed that
       that data in the respective columns contains valid values
   - \<ticker\> symbols are text of variable length
   - \<ticker\> symbols are grouped in consecutive rows for each unique symbol
   - A new \<ticker\> symbol signals processing for current symbol is complete
   - \<date\> values are in chronological order
       - First row for each \<ticker\> symbol contains open price for the year
       - Last row for each \<ticker\> symbols contains closing price for the year
   - \<date\> values have format YYYYMMDD. e.g. "20140131" for Jan 31, 2014
   - \<open\>, \<high\>, \<low\> and \<close\> are "float" values
   - \<volume\> values are of "long" type
   - All dates in the table are for the same year
   - A value of zero for first open price for a ticker symbol assumes there
     is no volume, hence no gain for the associated stock.
     
* **Assumption for output table :**
   - There is no other data in the spreadsheet beyond input table boundaries
   - Columns for output data will be cleared without loss of data
   - The first cell of the output table will be "I1"
   - The output can span columns "I" through "P"
   - The first row of output will contains headers as below
   
       | Ticker | Yearly Change | Percent Change | Total Volume |
       |---|---|---|---|
       
   - Ticker column will contain the \<ticker\> symbols for the stock
   - Yearly Change will contain "float" value for dollar amount
   - Percent Change will contain "float" value for percent value
   - Total Volume will contain "long" value for total shares traded


* **Assumptions for post processing output :**
   - The first cell of the output table will be "N1"
   - The output spans columns "N" through "P" worst case
   - The post processing output will be 3 x 4 area from "N1" containing
     labels and the values for least/most gain and highest volume

* **TODO :** Mitigate the assumptions from above :
   - Handle cases with zero for first opening price, check for subsequent
      non-zero price values and use for opening price
   - Support unordered table, ticker symbols not grouped, transactions
      not in chronalogical order
   - Support variable start column for input table, summary and post processSub StockAnalysis()