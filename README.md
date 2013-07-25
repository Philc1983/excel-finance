Excel Finance utilities
=============

Excel Finance is a set of free-to-use Excel VBA functions and utilities for Financial Analysis

## Functions

### functions/stock-yahoo.bas : GetTickerData
Excel Yahoo Stock Quote function takes (ticker, date, field) as input and outputs the result. 
It is useful to retrieve Close,Open,Volume,Low,High prices for any Yahoo stock ticker.
```
Usage
  =GetTickerData(TICKER, DATE, FIELD)
  Fields: Close,Open,Low,High,Volume
Example
  =GetTickerData(AAPL, "2013-1-7", "Close")
```


### Credits
<a href="http://fincluster.com/">Fincluster</a> : startup company focused on Financial Services and innovative Financial Platforms
