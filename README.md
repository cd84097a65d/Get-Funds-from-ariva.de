# Get-Funds-from-ariva.de
The table obtains important fund measures from ariva.de. 
Compatible to european funds. Some european funds cannot be obtained because they are not available from ariva.de.
The Excel should be set to german settings of numbers: "," is a decimal separator. 

The single important input parameter is a WKN (Wertpapierkennnummer, column C).
Other settings:
Favorites (Column D): Used to skip obtaining the funds data from ariva.de if not set ("") and "Ipdate only favorites" is checked.  

These colums are updated from ariva.de:
Country (Column E). The country is not always taken correctly. In this case it will stay empty.
Sector (Column F)
Benchmark (Column G)
Currency (Column H): Always EUR. I have never seen other currencies up to now. Therefore the sheet "Currencies", existed in previous version was removed.
URL (Column I): If the cell is empty, it will be automatically updated from ariva.de.
3m-5yrs (Columns J:N): Returns of the fund in a period of 3 months to 5 years. Can be empty, if the fund is younger.
Date (Column O): After obtaining of the measures it will be set to the today's date. It is also used to avoid getting of the fund measurements from the net (if the today's date is used). 
Price (Column P): Price in currency from (Column H).
Alpha (Column E): Alpha.
Beta (Column E): Beta.
Sharpe ratio (Column E): Sharpe ratio.

The sorting of categories can be done according to the selection in a list box (approx. at cell FA). The categories are listed in cells A59:A67. The VBA code search for the word "Sorting" above the list to give you more flexibility to add or remove funds.

Columns AB:AI contain the positions of the measures inside corresponding category (3m-Sharpe ratio). The "Sum" (Column AJ) contains the sum of the positions over all categories.

The measures in each category are highlighted according to the percentiles in rows 55-57. I prefere to use 85% percentile to highlight the values green and 15% percentiles to highlight them red.
