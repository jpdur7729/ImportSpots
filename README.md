# ImportSpots
Examples showing how to extract FX Spot Rates using public sources such as:
- ECB         (European Central Bank)
- ApiLayer    A free web service offering real time data for all currencies
- RijksBank   The Swedish Central Bank.
                 
Functional examples have been provided
1) ApiLayer   ==> Simple extract of all currencies against Euro
2) ECB-EUR    ==> Extract daily data against EUR
3) SEK-Based  ==> Mix approach to combine RiksBank to extract
SEK-based FXSpot for all najor currencies and then combine with
ApiLayer in order to get all the currencies against EUR

All examples provided useded:
- Powershell
- curl or wget to extract the data 
- FrontCmd in order to import a .csv file
                                         
Other providers not covered could be:
- Quandl                             
- Oanda                              
- Bank of England
