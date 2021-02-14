## Welcome to the further explanations to the table that obtains funds from ariva.de

The main idea of this page is to explain my way of investing to the funds. 

### Reasons
The main reason to develop the strategy was to create the investment strategy the works. Ariva.de was used as a source of the data because additionally to performance (the main part of strategy) it has also other financial measures like Sharpe ratio, alpha, beta and so on.

### Investment strategy
- Invest to the positions from the top third of the list of funds with best performence in the last 3 months.
- Do not touch bought positions for 3 months: do not sell them or do not buy the position again.
- If after three months the position is not inside the top third of the list of funds with best performence in the last 3 months, then it will be sold.
- 10% of return will be invested to gold. I am not a gold bug and do not believe that the gold will grow drastically, but this is a single possibility to save money at the time of the big crash.

### Proof of concept
The main question is: are the positions that are profitable today will give good returns in the future? I have tried to answer this question below:
- Sheet "Results", the graph at the left shows the dependency of performance in the future (6 months to 3 moths) as a function of performance in the past (1 year to 6 moths). The linear regression has a positive slope, which means that the profitable today positions will most probably stay profitable also in the future. 
- The graph at the right shows that with increasing Sharpe ratio, the positions will have increasing performance in 3 months.
- The average return of selected 17 (1/3 of 51) best positions in the past (best performance at 1 year to 6 months, seep cell P3) will give also a good averaged return in the future (6 months to 3 months, cell P4). Right now (14.02.2021) the average performance of best positions is even better than their performance in the past. This is connected with a fact that in a period from 1 year to 6 months there were a COVID-crysis and the positions were not recovered from the shock.

### TODO:
- Add parsing of other web pages which could contain the information that is missed at ariva.de.
- Make the tble indepencent on the representation of decimal point (now it is ",").
- May be check this strategies on shares.
