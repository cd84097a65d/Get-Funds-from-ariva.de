# The table that obtains funds from ariva.de
The table obtains important European funds measures from ariva.de. Some European funds cannot be obtained because they are not available from ariva.de. The Excel should be set to German settings of numbers: "," is a decimal separator.
Sometimes (pretty randomly) ariva.de returns undefined returns ("-"). In this case the "finanzen.net" is called.
The single important input parameter is a WKN (Wertpapierkennnummer, column C). 
Other settings: 
- Favorites (Column D): Used to skip obtaining the funds data from ariva.de if not set ("") and "Update only favorites" is checked.
These columns are updated from ariva.de: 
-Country (Column E). The country is not always taken correctly. In this case it will stay empty. 
- Sector (Column F) 
- Benchmark (Column G) 
- Currency (Column H): Underlying currency of the fund (since 10.02.2021)
- URL (Column I): If the cell is empty, it will be automatically updated from ariva.de. 
- 3m-5yrs (Columns J:N): Returns of the fund in a period of 3 months to 5 years. Can be empty if the fund is younger. 
- Date (Column O): After obtaining of the measures it will be set to the today's date. It is also used to avoid getting of the fund measurements from the net (if the today's date is used). 
- Price (Column P): Price in Euro. 
- Alpha (Column E): Alpha. 
- Beta (Column E): Beta. 
- Sharpe ratio (Column E): Sharpe ratio.

Additional measures from 10.02.2021:
- Volatility, 
- Tracking error, 
- Correlation,
- Skewness, 
- Kurtosis, 
- Sortino ratio, 
- Information ratio, 
- R^2, 
- Treynor ratio

The sorting of categories can be done according to the selection in a list box (at cell F1). The categories are listed in cells A57:A65. The VBA code searches for the word "Sorting" (A56) above the list to give you more flexibility to add or remove funds.
Columns AB:AI contains the positions of the measures inside corresponding category (3m-Sharpe ratio). The "Sum" (Column AJ) contains the sum of the positions over all categories. (since 10.02.2021: takes into account also the Sortino and Treynor ratios.)

The measures in each category are highlighted according to the percentiles in rows 53-55. I prefer to use 85% percentile to highlight the values green and 15% percentiles to highlight them red.
