## Welcome to the further explanations to the table that obtains funds from ariva.de

The main idea of this page is to explain my way of investing to the funds. Use it on your own risk, the author does not bear any risks for the future development of the stock exchanges. 

### Reasons
The main reason to develop the strategy was to create the investment strategy the works. Ariva.de was used as a source of the data because additionally to performance (the main part of strategy) it has also other financial measures like Sharpe ratio, alpha, beta and so on.

### Investment strategy
- Invest in positions from the top third of the list of funds with best performance in the last 3 months.
- Do not touch bought positions for 3 months: do not sell them or do not buy the position again.
- If after three months the position is not inside the top third of the list of funds with best performance in the last 3 months, then it will be sold.
- 10% of return will be invested to gold. I am not a “gold bug” and do not believe that the gold will grow drastically, but this is a single possibility to save money at the time of the big crash.

### Proof of concept
The main question is: are the positions that are profitable today will give good returns in the future? I have tried to answer this question below:
- Sheet "Results", the graph at the left shows the dependency of 3 months performance in the future (3 months to now and 6 months to 3 months) as a function of performance in the past (6 months to 3 months  and 1 year to 6 months). The linear regression has a positive slope, which means that the profitable today positions will most probably stay profitable also in the future. 
- My concept uses the investment in the positions from the top 1/3 of the positions with best performance over 3 months. The columns D and E show the positions with best performance (always over 3 months) 3 months ago and 6 months ago. The columns F and G show corresponding performance of these best positions now and 3 months ago. The performance now and 3 months ago is positive. 
- The corresponding average performance of the best positions 3 months ago and now and 6 months ago and 3 months ago is shown in cells P1:P2 and P4:P5. The conclusion: The positions with best performance in the past are also growing in the future!
- The graph at the right shows that with increasing Sharpe ratio (over one year), the positions will have increasing performance in 3 months.

### License:
No license. You are allowed to use, modify, distribute or sell the table or algorithms inside without any limitations. 

### TODO:
- Add parsing of other web pages which could contain the information that is missed at ariva.de.
- Make the table independent on the representation of decimal point (now it is ",").
- May be check this strategies on shares or at least compare this strategy with returns from the shares.
