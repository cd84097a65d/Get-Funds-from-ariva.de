# The table that obtains funds from ariva.de
The table obtains important European funds measures from ariva.de. The description of the goals of the table can be and proof of the concept can be found at the [GitHub Pages]( https://cd84097a65d.github.io/Get-Funds-from-ariva.de/).  

Some European funds cannot be obtained because they are not available from ariva.de. 

You have to install SeleniumBasic before you start to work with a table: https://florentbr.github.io/SeleniumBasic/
You must also download the ChromeDriver according to the version of Chrome browser from: 
https://sites.google.com/a/chromium.org/chromedriver/downloads
And copy the chromedriver.exe to the installation of selenium (in my case it is in C:\Users\%username%\AppData\Local\SeleniumBasic )

Sometimes (pretty randomly) ariva.de returns undefined performance ("-"). In this case the "finanzen.net" is called.
The single important input parameter is a WKN (Wertpapierkennnummer, column C). This parameter is used to find the corresponding web page at ariva.de or finanzen.net.
Other settings: 
- Favorites (Column D): Used to skip obtaining the funds data from ariva.de: if not set ("") and "Update only favorites" is checked.
- “Stock?” (Column F): Manually updated column that shows if the position available at stock exchange or only from the fund. I do not like to buy the positions which are available only from the fund. This means that you cannot sell this position within five minutes.

These columns are updated from ariva.de: 
- Position change (Column E): Change of the position since the last update. 
- Country (Column G). The country is not always taken correctly. In this case it will stay empty. 
- Sector (Column H) 
- Benchmark (Column I) 
- Currency (Column J): Underlying currency of the fund
- URL at ariva.de (Column K): If the cell is empty, it will be automatically updated from ariva.de. 
- URL at finanzen.net (Column L): If the cell is empty, it will be automatically updated from finanzen.net (if necessary). 
- 3m-5yrs (Columns M:Q): Returns of the fund in a period of 3 months to 5 years. Can be empty if the fund is younger. 
- Date (Column R): After obtaining of the measures it will be set to the today's date. It is also used to avoid getting of the fund measurements from the net (if the today's date is used). 
- Price (Column S): Price in Euro. 
- Alpha (Column T): Alpha. 
- Beta (Column U): Beta. 
- Sharpe ratio (Column V): Sharpe ratio.
- Volatility (column W), 
- Tracking error (column X), 
- Correlation (column Y),
- Skewness (column Z),
- Kurtosis (column AA),
- Sortino ratio (column AB),
- Information ratio (column AC),
- R^2 (column AD),
- Treynor ratio (column AE)

The sorting of categories can be done according to the selection in a list box (at cell F1). The categories are listed in cells A63:A74. The VBA code searches for the word "Sorting" (A62) above the list to give you more flexibility to add or remove funds.
Columns AQ:BA contains the positions of the measures inside corresponding category (3m-Sharpe ratio). The "Sum" (Column BB) contains the sum of the positions over all categories. (since 10.02.2021: takes into account also the Sortino and Treynor ratios.)

The measures in each category are highlighted according to the percentiles in rows 53-55. I prefer to use 85% percentile to highlight the values green and 15% percentiles to highlight them red.

The sheet “Results” represents a proof of the concept; for more information, read the [GitHub Pages]( https://cd84097a65d.github.io/Get-Funds-from-ariva.de/).
