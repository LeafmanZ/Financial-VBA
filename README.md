# FinancialVBA
Using VBA to analyze, calculate, or graph financial data interacting with excel.
Make sure you enable macros for each of the excel documents. The VBA code is tied as a module to its respective xlsm file.

CDF - Functions returns values from cumulative distribution function (CDF) and probability density function (PDF) using a Right Riemann Sum implementation

CRatio and HHI (2 functions)- 
Function 1: Calculates concentration ratios for each industry SIC then displays a histogram of all SICs, or graphs each industry and its respective concentration ratio. Only works for data collected for 1 given year. 
Function 2: Calculates HHI for each industry SIC then displays a histogram of all SICs, or graphs each industry and its respective HHI. Only works for data collected for 1 given year.

GBM - Subroutine (hotkey 'ctrl+g') Creates frequency distribution of predicted stock prices using general brownian motion.

Option Pricing Binomial - Uses binomial tree approach to calculate option price, delta, and vega. Can accept and factor in discrete, continuous dividends, and early exercise. Can do American or European.

Option Pricing Black Scholes - Uses Black Scholes approach to calculate option price, delta, and vega.  Can accept and factor in discrete, continuous dividends.

Option Pricing Implied Volatility - Uses Black Scholes approach to estimate implied volatility. Can accept and factor in discrete, continuous dividends. Can do American or European.

RW - Simulates random walk (Pretty pointless)

Option Pricing Monte Carlo - Uses Monte Carlo approach to calculate option price. Can factor in continuous dividends. Generates random walks for stock prices and calculate average option prices. Returns a 95% confidence interval of prices also.


Option Pricing Heston Model - Uses the Heston Model which generates random walks for volatility to update stock prices and calculate average option prices. Returns the price and 95% confidence interval of simulated european, up-and-out, and down-and-out options.

Feb1 Sharp and Sortino VBA - Calculates sharp and sortino variables.

Feb10 Deferred Payments VBA - Calculates deferred payments.

Feb15 Sheet Generation VBA - Generates new sheets.

Feb3 Ret Graph VBA - Generates graph of returns.

Feb8 Dist Graph VBA - Generates a distribution graph.

VBA on Powerpoint - Applying VBA to powerpoint.


