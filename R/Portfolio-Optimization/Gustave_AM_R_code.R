# LIBRARIES --------------------------------------------------------------------

install.packages("zoo")
install.packages("xts")
install.packages("PerformanceAnalytics")
install.packages("quantmod")
install.packages("ggplot2")
install.packages("tseries")
install.packages("fPortfolio")
install.packages("PortfolioAnalytics")
install.packages("ROI")
install.packages("ROI.plugin.quadprog")
install.packages("ROI.plugin.glpk")
install.packages("ROI.plugin.symphony")
install.packages("NMOF")
install.packages("corrplot")
install.packages("openxlsx")

library(zoo)
library(xts)
library(PerformanceAnalytics)
library(quantmod)
library(ggplot2)
library(tseries)
library(fPortfolio)
library(PortfolioAnalytics)
library(ROI)
library(ROI.plugin.quadprog)
library(ROI.plugin.glpk)
library(ROI.plugin.symphony)
library(NMOF)
library(corrplot)
library(openxlsx)

# END LIBRARIES ----------------------------------------------------------------


# FUNCTIONS --------------------------------------------------------------------

# Function: Extract daily returns from Yahoo Finance
Returns.Daily <- function(Tickers, StartDate, EndDate){
  PricesRead <- NULL
  # Loop to get prices of all tickers
  for (Ticker in Tickers){
    PricesRead <- cbind(PricesRead,getSymbols.yahoo(Ticker, from=StartDate,
                                                    to=EndDate, verbose=FALSE,
                                                    # Keep only the 6th column
                                                    # to get the Close Adjusted
                                                    auto.assign=FALSE)[,6])}
  Prices <- PricesRead[apply(PricesRead,1,function(x) all(!is.na(x))),]
  # Get the returns
  Returns <- na.omit(Return.calculate(Prices))
  # Clean the columns names
  colnames(Returns) <- gsub("\\.Adjusted", "", colnames(Returns))
  colnames(Returns) <- gsub("\\.", "-", colnames(Returns))
  return(Returns)}

# Function: Extract weekly returns from Yahoo Finance
Returns.Weekly <- function(Tickers, StartDate, EndDate){
  PricesRead <- NULL
  for (Ticker in Tickers){
    PricesRead <- cbind(PricesRead,getSymbols.yahoo(Ticker, from=StartDate,
                                                    to=EndDate, verbose=FALSE, 
                                                    auto.assign=FALSE)[,6])}
  Prices <- PricesRead[apply(PricesRead,1,function(x) all(!is.na(x))),]
  PricesWeekly <- to.weekly(Prices, OHLC = FALSE)
  Returns <- na.omit(Return.calculate(PricesWeekly))
  colnames(Returns) <- gsub("\\.Adjusted", "", colnames(Returns))
  colnames(Returns) <- gsub("\\.", "-", colnames(Returns))
  return(Returns)}

# Function: Extract monthly returns from Yahoo Finance
Returns.Monthly <- function(Tickers, StartDate, EndDate){
  PricesRead <- NULL
  for (Ticker in Tickers){
    PricesRead <- cbind(PricesRead,getSymbols.yahoo(Ticker, from=StartDate,
                                                    to=EndDate, verbose=FALSE, 
                                                    auto.assign=FALSE)[,6])}
  Prices <- PricesRead[apply(PricesRead,1,function(x) all(!is.na(x))),]
  PricesMonthly <- to.monthly(Prices, OHLC = FALSE)
  Returns <- na.omit(Return.calculate(PricesMonthly))
  colnames(Returns) <- gsub("\\.Adjusted", "", colnames(Returns))
  colnames(Returns) <- gsub("\\.", "-", colnames(Returns))
  return(Returns)}

# Function: compute Sharpe ratio of an equally-weighted portfolio from Tickers
Ptf.ew.Sr <- function(tickers) {
  if (!all(tickers %in% colnames(Returns))) {
    stop("Some tickers do not match columns names of Returns")}
  # Compute the portfolio return
  Ptf_ew_return <- mean(colMeans(Returns[,tickers]))
  # Compute the portfolio standard deviation
  Ptf_ew_sd <- sd(rowMeans(Returns[,tickers]))
  # Compute the portfolio Sharpe ratio
  Ptf_ew_sr <- Ptf_ew_return / Ptf_ew_sd # if rf = 0
  return(Ptf_ew_sr)}

# Function: Compute optimal weights for PTF_MINVAR
weights.minvar <- function(tickers, returns) {
  # Create a portfolio with chosen tickers
  initial_PTF_MINVAR <- portfolio.spec(assets = tickers)
  # Add weight constraints for each ticker
  initial_PTF_MINVAR <- add.constraint(initial_PTF_MINVAR, type = "box", 
                                          min = 0.01, max = 0.15)
  # Add objective
  initial_PTF_MINVAR <- add.objective(initial_PTF_MINVAR, type = "risk", 
                                         name = "var")
  # Find the min var portfolio
  PTF_MINVAR <- optimize.portfolio(returns, 
                                   initial_PTF_MINVAR, 
                                   optimize_method = "ROI")
  # Extract optimum weights
  weights_minvar <- PTF_MINVAR$weights
  return(weights_minvar)}

# Function: Compute optimal weights for PTF_MAXSR
weights.maxsr <- function(tickers, returns) {
  # Create a portfolio with chosen tickers
  initial_PTF_MAXSR <- portfolio.spec(assets=tickers)
  # Add weights constraints
  initial_PTF_MAXSR <- add.constraint(portfolio = initial_PTF_MAXSR, 
                                         type = "box", min = 0.01, max = 0.15)
  # add objectives
  initial_PTF_MAXSR <- add.objective(portfolio=initial_PTF_MAXSR, 
                                        type="return", name="mean")
  initial_PTF_MAXSR <- add.objective(portfolio=initial_PTF_MAXSR, 
                                        type="risk", name="StdDev")
  # Find the maxSR portfolio
  PTF_MAXSR <- optimize.portfolio(R=returns, 
                                  portfolio=initial_PTF_MAXSR, 
                                  optimize_method="ROI", 
                                  maxSR=TRUE, trace=TRUE)
  # Extract optimum weights
  weights_maxsr <- PTF_MAXSR$weights
  return(weights_maxsr)}

# END FUNCTIONS ----------------------------------------------------------------


### STEP 1 ---------------------------------------------------------------------
# -> Find best combination of 25 tickers from the 35 (already carefully chosen), 
# that maximize Sharpe ratio. For this step of "optimization", the portfolios 
# will be equally weighted. Moreover, we will do it in several steps (best 34 in 
# 35, then best 33 in 34... thanks to a loop, until we get the best 25), since 
# doing it in only 1 step is too time-consuming.
# ------------------------------------------------------------------------------

# Tickers of the 35 carefully chosen
Tickers_35 <- c("AAPL","MSFT","NVDA","GOOGL","AMD","V","ADBE","META","NFLX",
                "CRM","TXN","INTC","CSCO","ORCL","ACN","LLY","JNJ","UNH",
                "MRK","ABBV","AMZN","TSLA","MCD","DIS","BRK-B","KO","WMT",
                "CAT","UNP","BA","JPM","BAC","NEE","LIN","PLD")


# 2021
# Start by taking the 35
best_25_y21 = Tickers_35
# Get returns from 2 previous years
Returns <- Returns.Daily(best_25_y21, "2019-01-01", "2020-12-31")
# Convert Returns into a DataFrame to prevent not matching colNames issues with 
# my Ptf.ew.Sr function
Returns <- as.data.frame(Returns)
# loop until we get the best 25
for (n in seq(length(best_25_y21)-1, 25, by = -1)){
  combinations <- combn(best_25_y21, n, simplify = FALSE)
  sharpe_ratios <- sapply(combinations, Ptf.ew.Sr)
  best_25_y21 <- combinations[[which.max(sharpe_ratios)]]}
best_25_y21


# 2022 
# (same logic)
best_25_y22 = Tickers_35
Returns <- Returns.Daily(best_25_y22, "2020-01-01", "2021-12-31")
Returns <- as.data.frame(Returns)
for (n in seq(length(best_25_y22)-1, 25, by = -1)){
  combinations <- combn(best_25_y22, n, simplify = FALSE)
  sharpe_ratios <- sapply(combinations, Ptf.ew.Sr)
  best_25_y22 <- combinations[[which.max(sharpe_ratios)]]}
best_25_y22


# 2023
best_25_y23 = Tickers_35
Returns <- Returns.Daily(best_25_y23, "2021-01-01", "2022-12-31")
Returns <- as.data.frame(Returns)
for (n in seq(length(best_25_y23)-1, 25, by = -1)){
  combinations <- combn(best_25_y23, n, simplify = FALSE)
  sharpe_ratios <- sapply(combinations, Ptf.ew.Sr)
  best_25_y23 <- combinations[[which.max(sharpe_ratios)]]}
best_25_y23

# END STEP 1 -------------------------------------------------------------------




### STEP 2 ---------------------------------------------------------------------
# -> Build minvar and maxsr portfolios for 2021, 2022 and 2023
# ------------------------------------------------------------------------------

## 1 - 2021 (Find optimum weights for 2021, based on 2019 and 2020 returns)

# Take the best 25 tickers for 2021
Tickers_y21 <- best_25_y21

# Returns of 2019 and 2020 with best 25 tickers
Returns_dly_19_20 <- Returns.Daily(Tickers_y21, "2019-01-01", "2020-12-31")

# a) PTF_MINVAR (to get optimum weights for the min var ptf, we use our 
# weights.minvar function that we built - cf. FUNCTIONS part)
weights_minvar_21 <- weights.minvar(Tickers_y21, Returns_dly_19_20)

# b) PTF_MAXSR (same logic)
weights_maxsr_21 <- weights.maxsr(Tickers_y21, Returns_dly_19_20)

# Data Quality Check: We checked that our functions weights.minvar and 
# weights.maxsr work correctly by solving the weights thanks to the ready-to-use
# functions below that come from another library, and we found the exact same 
# weights.
weights_minvar_21_TEST <- minvar(cov(Returns_dly_19_20), wmin = 0.01, 
                                 wmax = 0.15, method = "qp")
weights_maxsr_21_TEST <- maxSharpe(colMeans(Returns_dly_19_20), 
                                   cov(Returns_dly_19_20), wmin = 0.01, 
                                   wmax = 0.15, method = "qp")
# Data Quality Check: OK!


## 2 - 2022 (Find optimum weights for 2022, based on 2020 and 2021 returns)

# Take the best 25 tickers for 2022
Tickers_y22 <- best_25_y22

# Returns of 2020 and 2021 with best 25 tickers
Returns_dly_20_21 <- Returns.Daily(Tickers_y22, "2020-01-01", "2021-12-31")

# a) PTF_MINVAR
weights_minvar_22 <- weights.minvar(Tickers_y22, Returns_dly_20_21)

# b) PTF_MAXSR
weights_maxsr_22 <- weights.maxsr(Tickers_y22, Returns_dly_20_21)


## 3 - 2023 (Find optimum weights for 2023, based on 2021 and 2022 returns)

# Take the best 25 tickers for 2023
Tickers_y23 <- best_25_y23

# Returns of 2021 and 2022 with best 25 tickers
Returns_dly_21_22 <- Returns.Daily(Tickers_y23, "2021-01-01", "2022-12-31")

# a) PTF_MINVAR
weights_minvar_23 <- weights.minvar(Tickers_y23, Returns_dly_21_22)

# b) PTF_MAXSR
weights_maxsr_23 <- weights.maxsr(Tickers_y23, Returns_dly_21_22)

# END STEP 2 -------------------------------------------------------------------




### STEP 3 ---------------------------------------------------------------------
# -> Compute all the required REPORTING
# ------------------------------------------------------------------------------

## 1 - 2021

# a) Benchmark : S&P500 index
SP500_returns_dly_21 <- Returns.Daily("^GSPC", "2021-01-01", "2021-12-31")
SP500_Ret_21 <- mean(SP500_returns_dly_21)*252
SP500_SD_21 <- sd(SP500_returns_dly_21)*sqrt(252)
SP500_Sr_21 <- SP500_Ret_21 / SP500_SD_21
SP500_maxDD_21 <- maxDrawdown(SP500_returns_dly_21)

# b) PTF_MINVAR
Returns_dly_21 <- Returns.Daily(Tickers_y21, "2021-01-01", "2021-12-31")
stocks_mu_21 <- colMeans(Returns_dly_21)

PTF_MINVAR_Ret_21 <- (t(weights_minvar_21) %*% stocks_mu_21)*252
PTF_MINVAR_SD_21 <- sqrt(t(weights_minvar_21) %*% cov(Returns_dly_21) %*% 
                           weights_minvar_21)*sqrt(252)
PTF_MINVAR_Sr_21 <- PTF_MINVAR_Ret_21 / PTF_MINVAR_SD_21 # rf = 0

# c) PTF_MAXSR
Returns_dly_20 <- Returns.Daily(Tickers_y21, "2020-01-01", "2020-12-31")
Returns_dly_21 <- Returns.Daily(Tickers_y21, "2021-01-01", "2021-12-31")
stocks_mu_21 <- colMeans(Returns_dly_21)

PTF_MAXSR_Ret_21 <- (t(weights_maxsr_21) %*% stocks_mu_21)*252
PTF_MAXSR_SD_21 <- sqrt(t(weights_maxsr_21) %*% cov(Returns_dly_21) %*% 
                          weights_maxsr_21) * sqrt(252)
PTF_MAXSR_Sr_21 <- PTF_MAXSR_Ret_21 / PTF_MAXSR_SD_21

# d) FUND_GAM
weights_ew_21 <- setNames(rep(1/25, times = 25), Tickers_y21)

Returns_dly_21 <- Returns.Daily(Tickers_y21, "2021-01-01", "2021-12-31")
stocks_mu_21 <- colMeans(Returns_dly_21)

FUND_GAM_Ret_21 <- (t(weights_ew_21) %*% stocks_mu_21)*252
FUND_GAM_SD_21 <- sqrt(t(weights_ew_21) %*% cov(Returns_dly_21) %*% 
                        weights_ew_21) * sqrt(252)
FUND_GAM_Sr_21 <- FUND_GAM_Ret_21 / FUND_GAM_SD_21

# Beta (with monthly returns)
Returns_mly_21 <- Returns.Monthly(Tickers_y21, "2021-01-01", "2021-12-31")
FUND_GAM_returns_mly_21 <- Returns_mly_21 %*% weights_ew_21
SP500_returns_mly_21 <- Returns.Monthly("^GSPC", "2021-01-01", "2021-12-31")
FUND_GAM_Beta_21 <- cov(FUND_GAM_returns_mly_21, SP500_returns_mly_21) / 
                    var(SP500_returns_mly_21)

# Treynor ratio
FUND_GAM_Treynor_21 <- FUND_GAM_Ret_21 / FUND_GAM_Beta_21 # rf = 0

# SRRI of GAM FUND (by calculate average annual volatility over the last 5 years 
# using weekly returns)
Returns_wkly_17to21 <- Returns.Weekly(best_25_y21, StartDate = "2017-01-01", 
                                      EndDate = "2021-12-31")
FUND_GAM_SD_17to21 <- sqrt(t(weights_ew_21) %*% cov(Returns_wkly_17to21) %*% 
                            weights_ew_21) * sqrt(52)
# FUND_GAM_SD_17to21 ~ 18%, between 15% and 25%, so FUND_GAM_SRRI_21 = 6 
FUND_GAM_SRRI_21 <- 6

# Alpha (with rf = 0)
FUND_GAM_Alpha_21 <- FUND_GAM_Ret_21 - (FUND_GAM_Beta_21 * SP500_Ret_21)

# Maxdrawdown
FUND_GAM_returns_dly_21 <- Return.portfolio(Returns_dly_21, weights_ew_21)
FUND_GAM_maxDD_21 <- maxDrawdown(FUND_GAM_returns_dly_21)

# Tracking error
FUND_GC_trackerr_21 <- sd(FUND_GAM_returns_dly_21 - SP500_returns_dly_21)

# Wealth index 2021 of FUND_GAM and S&P 500
FUND_GAM_index_21 <- Return.portfolio(Returns_dly_21,wealth.index = T)
SP500_index_21 <- Return.portfolio(SP500_returns_dly_21,wealth.index = T)
plot(index(FUND_GAM_index_21),FUND_GAM_index_21, type = "l", col = "blue", 
     lty = 1, xlab = "Time", ylab = "Wealth Index", xaxt="n", 
     main = "Croissance de 1 000 000 EUR - Année civile")
lines(index(FUND_GAM_index_21), SP500_index_21, col = 'orange')
legend(x = "topleft",legend = c("Gustave Equ. Tech. Sustainable US", 
                                "Indice de référence : S&P 500 Index"),
       lty = c(1, 1),col = c("blue", "orange"),lwd = 2)
axis(1, at = index(FUND_GAM_index_21)[seq(1, length(index(FUND_GAM_index_21)), 
                                           by = 21)], 
     labels = format(index(FUND_GAM_index_21)
                     [seq(1, length(index(FUND_GAM_index_21)), by = 21)], "%b"), 
     cex.axis = 0.7, las = 2)


## 2 - 2022

# a) Benchmark
SP500_returns_dly_22 <- Returns.Daily("^GSPC", "2022-01-01", "2022-12-31")
SP500_Ret_22 <- mean(SP500_returns_dly_22)*252
SP500_SD_22 <- sd(SP500_returns_dly_22)*sqrt(252)
SP500_Sr_22 <- SP500_Ret_22 / SP500_SD_22

# b) PTF_MINVAR
Returns_dly_22 <- Returns.Daily(Tickers_y22, "2022-01-01", "2022-12-31")
stocks_mu_22 <- colMeans(Returns_dly_22)

PTF_MINVAR_Ret_22 <- (t(weights_minvar_22) %*% stocks_mu_22)*252
PTF_MINVAR_SD_22 <- sqrt(t(weights_minvar_22) %*% cov(Returns_dly_22) %*% 
                           weights_minvar_22)*sqrt(252)
PTF_MINVAR_Sr_22 <- PTF_MINVAR_Ret_22 / PTF_MINVAR_SD_22

# c) PTF_MAXSR
Returns_dly_22 <- Returns.Daily(Tickers_y22, "2022-01-01", "2022-12-31")
stocks_mu_22 <- colMeans(Returns_dly_22)

PTF_MAXSR_Ret_22 <- (t(weights_maxsr_22) %*% stocks_mu_22)*252
PTF_MAXSR_SD_22 <- sqrt(t(weights_maxsr_22) %*% cov(Returns_dly_22) 
                        %*% weights_maxsr_22) * sqrt(252)
PTF_MAXSR_Sr_22 <- PTF_MAXSR_Ret_22 / PTF_MAXSR_SD_22

# d) FUND_GAM
weights_ew_22 <- setNames(rep(1/25, times = 25), Tickers_y22)

Returns_dly_22 <- Returns.Daily(Tickers_y22, "2022-01-01", "2022-12-31")
stocks_mu_22 <- colMeans(Returns_dly_22)

FUND_GAM_Ret_22 <- (t(weights_ew_22) %*% stocks_mu_22)*252
FUND_GAM_SD_22 <- sqrt(t(weights_ew_22) %*% cov(Returns_dly_22) 
                      %*% weights_ew_22) * sqrt(252)
FUND_GAM_Sr_22 <- FUND_GAM_Ret_22 / FUND_GAM_SD_22

Returns_mly_22 <- Returns.Monthly(Tickers_y22, "2022-01-01", "2022-12-31")
FUND_GAM_returns_mly_22 <- Returns_mly_22 %*% weights_ew_22
SP500_returns_mly_22 <- Returns.Monthly("^GSPC", "2022-01-01", "2022-12-31")
FUND_GAM_Beta_22 <- cov(FUND_GAM_returns_mly_22, SP500_returns_mly_22) / 
                    var(SP500_returns_mly_22)

FUND_GAM_Treynor_22 <- FUND_GAM_Ret_22 / FUND_GAM_Beta_22

Returns_wkly_18to22 <- Returns.Weekly(best_25_y22, StartDate = "2018-01-01", 
                                      EndDate = "2022-12-31")
FUND_GAM_SD_18to22 <- sqrt(t(weights_ew_22) %*% cov(Returns_wkly_18to22) %*% 
                            weights_ew_22) * sqrt(52)
FUND_GAM_SRRI_22 <- 6 # FUND_GAM_SD_18to22 ~ 22%

FUND_GAM_Alpha_22 <- FUND_GAM_Ret_22 - (FUND_GAM_Beta_22 * SP500_Ret_22)

FUND_GAM_returns_dly_22 <- Return.portfolio(Returns_dly_22, weights_ew_22)
FUND_GAM_maxDD_22 <- maxDrawdown(FUND_GAM_returns_dly_22)

FUND_GC_trackerr_22 <- sd(FUND_GAM_returns_dly_22 - SP500_returns_dly_22)

FUND_GAM_index_22 <- Return.portfolio(Returns_dly_22,wealth.index = T)
SP500_index_22 <- Return.portfolio(SP500_returns_dly_22,wealth.index = T)
plot(index(FUND_GAM_index_22),FUND_GAM_index_22, type = "l", col = "blue", 
     lty = 1, xlab = "Time", ylab = "Wealth Index", xaxt="n", 
     main = "Croissance de 1 000 000 EUR - Année civile")
lines(index(FUND_GAM_index_22), SP500_index_22, col = 'orange')
legend(x = "topleft",legend = c("Gustave Equ. Tech. Sustainable US", 
                                "Indice de référence : S&P 500 Index"),
       lty = c(1, 1),col = c("blue", "orange"),lwd = 2)
axis(1, at = index(FUND_GAM_index_22)[seq(1, length(index(FUND_GAM_index_22)), 
                                          by = 21)], 
     labels = format(index(FUND_GAM_index_22)
                     [seq(1, length(index(FUND_GAM_index_22)), by = 21)], "%b"), 
     cex.axis = 0.7, las = 2)


## 3 - 2023

# a) Benchmark
SP500_returns_dly_23 <- Returns.Daily("^GSPC", "2023-01-01", "2023-12-31")
SP500_Ret_23 <- mean(SP500_returns_dly_23)*252
SP500_SD_23 <- sd(SP500_returns_dly_23)*sqrt(252)
SP500_Sr_23 <- SP500_Ret_23 / SP500_SD_23

# b) PTF_MINVAR
Returns_dly_23 <- Returns.Daily(Tickers_y23, "2023-01-01", "2023-12-31")
stocks_mu_23 <- colMeans(Returns_dly_23)

PTF_MINVAR_Ret_23 <- (t(weights_minvar_23) %*% stocks_mu_23)*252
PTF_MINVAR_SD_23 <- sqrt(t(weights_minvar_23) %*% cov(Returns_dly_23) %*% 
                           weights_minvar_23)*sqrt(252)
PTF_MINVAR_Sr_23 <- PTF_MINVAR_Ret_23 / PTF_MINVAR_SD_23

# c) PTF_MAXSR
Returns_dly_23 <- Returns.Daily(Tickers_y23, "2023-01-01", "2023-12-31")
stocks_mu_23 <- colMeans(Returns_dly_23)

PTF_MAXSR_Ret_23 <- (t(weights_maxsr_23) %*% stocks_mu_23)*252
PTF_MAXSR_SD_23 <- sqrt(t(weights_maxsr_23) %*% cov(Returns_dly_23) %*% 
                          weights_maxsr_23) * sqrt(252)
PTF_MAXSR_Sr_23 <- PTF_MAXSR_Ret_23 / PTF_MAXSR_SD_23

# d) FUND_GAM
weights_ew_23 <- setNames(rep(1/25, times = 25), Tickers_y23)

Returns_dly_23 <- Returns.Daily(Tickers_y23, "2023-01-01", "2023-12-31")
stocks_mu_23 <- colMeans(Returns_dly_23)

FUND_GAM_Ret_23 <- (t(weights_ew_23) %*% stocks_mu_23)*252
FUND_GAM_SD_23 <- sqrt(t(weights_ew_23) %*% cov(Returns_dly_23) %*% 
                        weights_ew_23) * sqrt(252)
FUND_GAM_Sr_23 <- FUND_GAM_Ret_23 / FUND_GAM_SD_23

Returns_mly_23 <- Returns.Monthly(Tickers_y23, "2023-01-01", "2023-12-31")
FUND_GAM_returns_mly_23 <- Returns_mly_23 %*% weights_ew_23
SP500_returns_mly_23 <- Returns.Monthly("^GSPC", "2023-01-01", "2023-12-31")
FUND_GAM_Beta_23 <- cov(FUND_GAM_returns_mly_23, SP500_returns_mly_23) / 
                    var(SP500_returns_mly_23)

FUND_GAM_Treynor_23 <- FUND_GAM_Ret_23 / FUND_GAM_Beta_23

Returns_wkly_19to23 <- Returns.Weekly(best_25_y23, StartDate = "2019-01-01", 
                                      EndDate = "2023-12-31")
FUND_GAM_SD_19to23 <- sqrt(t(weights_ew_23) %*% cov(Returns_wkly_19to23) %*% 
                            weights_ew_23) * sqrt(52)
FUND_GAM_SRRI_23 <- 6 # FUND_GAM_SD_18to23 ~ 21%

FUND_GAM_Alpha_23 <- FUND_GAM_Ret_23 - (FUND_GAM_Beta_23 * SP500_Ret_23)

FUND_GAM_returns_dly_23 <- Return.portfolio(Returns_dly_23, weights_ew_23)
FUND_GAM_maxDD_23 <- maxDrawdown(FUND_GAM_returns_dly_23)

FUND_GC_trackerr_23 <- sd(FUND_GAM_returns_dly_23 - SP500_returns_dly_23)

FUND_GAM_index_23 <- Return.portfolio(Returns_dly_23,wealth.index = T)
SP500_index_23 <- Return.portfolio(SP500_returns_dly_23,wealth.index = T)
plot(index(FUND_GAM_index_23),FUND_GAM_index_23, type = "l", col = "blue", 
     lty = 1, xlab = "Time", ylab = "Wealth Index", xaxt="n", 
     main = "Croissance de 1 000 000 EUR - Année civile")
lines(index(FUND_GAM_index_23), SP500_index_23, col = 'orange')
legend(x = "topleft",legend = c("Gustave Equ. Tech. Sustainable US", 
                                "Indice de référence : S&P 500 Index"),
       lty = c(1, 1),col = c("blue", "orange"),lwd = 2)
axis(1, at = index(FUND_GAM_index_23)[seq(1, length(index(FUND_GAM_index_23)), 
                                          by = 21)], 
     labels = format(index(FUND_GAM_index_23)
                     [seq(1, length(index(FUND_GAM_index_23)), by = 21)], "%b"), 
     cex.axis = 0.7, las = 2)


## 4 - all years

# Correlation matrix of the 35 tickers
Returns <- Returns.Daily(Tickers_35, "2021-01-01", "2023-12-31")
correlation_matrix <- cor(Returns)

# Plot the correlation matrix
corrplot(correlation_matrix, method = "color", type = "upper", order = "hclust", 
         addCoef.col = "black", tl.col = "black", tl.srt = 45, tl.cex = 0.6,
         number.cex = 0.5,)

# END STEP 3 -------------------------------------------------------------------




### STEP 4 ---------------------------------------------------------------------
# -> Extract reporting data into an excel workbook
# ------------------------------------------------------------------------------

## Convert data into a DataFrame

# 2021
# S&P500_2021
SP500_Ret_21_DF <- as.data.frame(SP500_Ret_21)
SP500_SD_21_DF <- as.data.frame(SP500_SD_21)
SP500_Sr_21_DF <- as.data.frame(SP500_Sr_21)

# PTF_MINVAR_2021
weights_minvar_21_DF <- as.data.frame(weights_minvar_21)
PTF_MINVAR_Ret_21_DF <- setNames(data.frame(PTF_MINVAR_Ret_21), 
                                 "PTF_MINVAR_Ret_21")
PTF_MINVAR_SD_21_DF <- setNames(as.data.frame(PTF_MINVAR_SD_21), 
                                "PTF_MINVAR_SD_21")
PTF_MINVAR_Sr_21_DF <- setNames(as.data.frame(PTF_MINVAR_Sr_21), 
                                "PTF_MINVAR_Sr_21")

# PTF_MAXSR_2021
weights_maxsr_21_DF <- as.data.frame(weights_maxsr_21)
PTF_MAXSR_Ret_21_DF <- setNames(data.frame(PTF_MAXSR_Ret_21), 
                                "PTF_MAXSR_Ret_21")
PTF_MAXSR_SD_21_DF <- setNames(data.frame(PTF_MAXSR_SD_21), "PTF_MAXSR_SD_21")
PTF_MAXSR_Sr_21_DF <- setNames(data.frame(PTF_MAXSR_Sr_21), "PTF_MAXSR_Sr_21")

# FUND_GAM_2021
weights_ew_21_DF <- as.data.frame(weights_ew_21)
FUND_GAM_Ret_21_DF <- setNames(data.frame(FUND_GAM_Ret_21), "FUND_GAM_Ret_21")
FUND_GAM_SD_21_DF <- setNames(data.frame(FUND_GAM_SD_21), "FUND_GAM_SD_21")
FUND_GAM_Sr_21_DF <- setNames(data.frame(FUND_GAM_Sr_21), "FUND_GAM_Sr_21")

FUND_GAM_Beta_21_DF <- setNames(data.frame(FUND_GAM_Beta_21), 
                                "FUND_GAM_Beta_21")
FUND_GAM_Treynor_21_DF <- setNames(data.frame(FUND_GAM_Treynor_21), 
                                   "FUND_GAM_Treynor_21")
FUND_GAM_SRRI_21_DF <- setNames(data.frame(FUND_GAM_SRRI_21), 
                                "FUND_GAM_SRRI_21")
FUND_GAM_Alpha_21_DF <- setNames(data.frame(FUND_GAM_Alpha_21), 
                                 "FUND_GAM_Alpha_21")

# 2022
# S&P500_2022
SP500_Ret_22_DF <- as.data.frame(SP500_Ret_22)
SP500_SD_22_DF <- as.data.frame(SP500_SD_22)
SP500_Sr_22_DF <- as.data.frame(SP500_Sr_22)

# PTF_MINVAR_2022
weights_minvar_22_DF <- as.data.frame(weights_minvar_22)
PTF_MINVAR_Ret_22_DF <- setNames(data.frame(PTF_MINVAR_Ret_22), 
                                 "PTF_MINVAR_Ret_22")
PTF_MINVAR_SD_22_DF <- setNames(as.data.frame(PTF_MINVAR_SD_22), 
                                "PTF_MINVAR_SD_22")
PTF_MINVAR_Sr_22_DF <- setNames(as.data.frame(PTF_MINVAR_Sr_22), 
                                "PTF_MINVAR_Sr_22")

# PTF_MAXSR_2022
weights_maxsr_22_DF <- as.data.frame(weights_maxsr_22)
PTF_MAXSR_Ret_22_DF <- setNames(data.frame(PTF_MAXSR_Ret_22), 
                                "PTF_MAXSR_Ret_22")
PTF_MAXSR_SD_22_DF <- setNames(data.frame(PTF_MAXSR_SD_22), "PTF_MAXSR_SD_22")
PTF_MAXSR_Sr_22_DF <- setNames(data.frame(PTF_MAXSR_Sr_22), "PTF_MAXSR_Sr_22")

# FUND_GAM_2022
weights_ew_22_DF <- as.data.frame(weights_ew_22)
FUND_GAM_Ret_22_DF <- setNames(data.frame(FUND_GAM_Ret_22), "FUND_GAM_Ret_22")
FUND_GAM_SD_22_DF <- setNames(data.frame(FUND_GAM_SD_22), "FUND_GAM_SD_22")
FUND_GAM_Sr_22_DF <- setNames(data.frame(FUND_GAM_Sr_22), "FUND_GAM_Sr_22")

FUND_GAM_Beta_22_DF <- setNames(data.frame(FUND_GAM_Beta_22), 
                                "FUND_GAM_Beta_22")
FUND_GAM_Treynor_22_DF <- setNames(data.frame(FUND_GAM_Treynor_22), 
                                   "FUND_GAM_Treynor_22")
FUND_GAM_SRRI_22_DF <- setNames(data.frame(FUND_GAM_SRRI_22), 
                                "FUND_GAM_SRRI_22")
FUND_GAM_Alpha_22_DF <- setNames(data.frame(FUND_GAM_Alpha_22), 
                                 "FUND_GAM_Alpha_22")

# 2023
# S&P500_2023
SP500_Ret_23_DF <- as.data.frame(SP500_Ret_23)
SP500_SD_23_DF <- as.data.frame(SP500_SD_23)
SP500_Sr_23_DF <- as.data.frame(SP500_Sr_23)

# PTF_MINVAR_2023
weights_minvar_23_DF <- as.data.frame(weights_minvar_23)
PTF_MINVAR_Ret_23_DF <- setNames(data.frame(PTF_MINVAR_Ret_23), 
                                 "PTF_MINVAR_Ret_23")
PTF_MINVAR_SD_23_DF <- setNames(as.data.frame(PTF_MINVAR_SD_23), 
                                "PTF_MINVAR_SD_23")
PTF_MINVAR_Sr_23_DF <- setNames(as.data.frame(PTF_MINVAR_Sr_23), 
                                "PTF_MINVAR_Sr_23")

# PTF_MAXSR_2023
weights_maxsr_23_DF <- as.data.frame(weights_maxsr_23)
PTF_MAXSR_Ret_23_DF <- setNames(data.frame(PTF_MAXSR_Ret_23), 
                                "PTF_MAXSR_Ret_23")
PTF_MAXSR_SD_23_DF <- setNames(data.frame(PTF_MAXSR_SD_23), "PTF_MAXSR_SD_23")
PTF_MAXSR_Sr_23_DF <- setNames(data.frame(PTF_MAXSR_Sr_23), "PTF_MAXSR_Sr_23")

# FUND_GAM_2023
weights_ew_23_DF <- as.data.frame(weights_ew_23)
FUND_GAM_Ret_23_DF <- setNames(data.frame(FUND_GAM_Ret_23), "FUND_GAM_Ret_23")
FUND_GAM_SD_23_DF <- setNames(data.frame(FUND_GAM_SD_23), "FUND_GAM_SD_23")
FUND_GAM_Sr_23_DF <- setNames(data.frame(FUND_GAM_Sr_23), "FUND_GAM_Sr_23")

FUND_GAM_Beta_23_DF <- setNames(data.frame(FUND_GAM_Beta_23), 
                                "FUND_GAM_Beta_23")
FUND_GAM_Treynor_23_DF <- setNames(data.frame(FUND_GAM_Treynor_23), 
                                   "FUND_GAM_Treynor_23")
FUND_GAM_SRRI_23_DF <- setNames(data.frame(FUND_GAM_SRRI_23), 
                                "FUND_GAM_SRRI_23")
FUND_GAM_Alpha_23_DF <- setNames(data.frame(FUND_GAM_Alpha_23), 
                                 "FUND_GAM_Alpha_23")


## Export data into an xlsx
wb <- createWorkbook()


# 2021
addWorksheet(wb, "2021")

# S&P500
writeData(wb, sheet = "2021", 
          "BENCHMARK",
          startRow = 2, startCol = 2)
writeData(wb, sheet = "2021", 
          SP500_Ret_21_DF,
          startRow = 4, startCol = 2)
writeData(wb, sheet = "2021", 
          SP500_SD_21_DF,
          startRow = 7, startCol = 2)
writeData(wb, sheet = "2021", 
          SP500_Sr_21_DF,
          startRow = 10, startCol = 2)

# PTF_MINVAR
writeData(wb, sheet = "2021", 
          "PTF MIN VAR",
          startRow = 2, startCol = 7)
writeData(wb, sheet = "2021", 
          weights_minvar_21_DF,
          startRow = 4, startCol = 5,
          rowNames = TRUE)
writeData(wb, sheet = "2021", 
          PTF_MINVAR_Ret_21_DF,
          startRow = 4, startCol = 8)
writeData(wb, sheet = "2021", 
          PTF_MINVAR_SD_21_DF,
          startRow = 7, startCol = 8)
writeData(wb, sheet = "2021", 
          PTF_MINVAR_Sr_21_DF,
          startRow = 10, startCol = 8)

# PTF_MAXSR
writeData(wb, sheet = "2021", 
          "PTF MAX SR",
          startRow = 2, startCol = 13)
writeData(wb, sheet = "2021", 
          weights_maxsr_21_DF,
          startRow = 4, startCol = 11,
          rowNames = TRUE)
writeData(wb, sheet = "2021", 
          PTF_MAXSR_Ret_21_DF,
          startRow = 4, startCol = 14)
writeData(wb, sheet = "2021", 
          PTF_MAXSR_SD_21_DF,
          startRow = 7, startCol = 14)
writeData(wb, sheet = "2021", 
          PTF_MAXSR_Sr_21_DF,
          startRow = 10, startCol = 14)

# FUND_GAM
writeData(wb, sheet = "2021", 
          "FOND GAM",
          startRow = 2, startCol = 19)
writeData(wb, sheet = "2021", 
          weights_ew_21_DF,
          startRow = 4, startCol = 17,
          rowNames = TRUE)
writeData(wb, sheet = "2021", 
          FUND_GAM_Ret_21_DF,
          startRow = 4, startCol = 20)
writeData(wb, sheet = "2021", 
          FUND_GAM_SD_21_DF,
          startRow = 7, startCol = 20)
writeData(wb, sheet = "2021", 
          FUND_GAM_Sr_21_DF,
          startRow = 10, startCol = 20)
writeData(wb, sheet = "2021", 
          FUND_GAM_Beta_21_DF,
          startRow = 13, startCol = 20)
writeData(wb, sheet = "2021", 
          FUND_GAM_Treynor_21_DF,
          startRow = 16, startCol = 20)
writeData(wb, sheet = "2021", 
          FUND_GAM_Alpha_21_DF,
          startRow = 19, startCol = 20)
writeData(wb, sheet = "2021", 
          FUND_GAM_SRRI_21_DF,
          startRow = 22, startCol = 20)

setColWidths(wb, "2021", cols = 1:30, widths = "auto")


# 2022
addWorksheet(wb, "2022")

# S&P500
writeData(wb, sheet = "2022", 
          "BENCHMARK",
          startRow = 2, startCol = 2)
writeData(wb, sheet = "2022", 
          SP500_Ret_22_DF,
          startRow = 4, startCol = 2)
writeData(wb, sheet = "2022", 
          SP500_SD_22_DF,
          startRow = 7, startCol = 2)
writeData(wb, sheet = "2022", 
          SP500_Sr_22_DF,
          startRow = 10, startCol = 2)

# PTF_MINVAR
writeData(wb, sheet = "2022", 
          "PTF MIN VAR",
          startRow = 2, startCol = 7)
writeData(wb, sheet = "2022", 
          weights_minvar_22_DF,
          startRow = 4, startCol = 5,
          rowNames = TRUE)
writeData(wb, sheet = "2022", 
          PTF_MINVAR_Ret_22_DF,
          startRow = 4, startCol = 8)
writeData(wb, sheet = "2022", 
          PTF_MINVAR_SD_22_DF,
          startRow = 7, startCol = 8)
writeData(wb, sheet = "2022", 
          PTF_MINVAR_Sr_22_DF,
          startRow = 10, startCol = 8)

# PTF_MAXSR
writeData(wb, sheet = "2022", 
          "PTF MAX SR",
          startRow = 2, startCol = 13)
writeData(wb, sheet = "2022", 
          weights_maxsr_22_DF,
          startRow = 4, startCol = 11,
          rowNames = TRUE)
writeData(wb, sheet = "2022", 
          PTF_MAXSR_Ret_22_DF,
          startRow = 4, startCol = 14)
writeData(wb, sheet = "2022", 
          PTF_MAXSR_SD_22_DF,
          startRow = 7, startCol = 14)
writeData(wb, sheet = "2022", 
          PTF_MAXSR_Sr_22_DF,
          startRow = 10, startCol = 14)

# FUND_GAM
writeData(wb, sheet = "2022", 
          "FOND GAM",
          startRow = 2, startCol = 19)
writeData(wb, sheet = "2022", 
          weights_ew_22_DF,
          startRow = 4, startCol = 17,
          rowNames = TRUE)
writeData(wb, sheet = "2022", 
          FUND_GAM_Ret_22_DF,
          startRow = 4, startCol = 20)
writeData(wb, sheet = "2022", 
          FUND_GAM_SD_22_DF,
          startRow = 7, startCol = 20)
writeData(wb, sheet = "2022", 
          FUND_GAM_Sr_22_DF,
          startRow = 10, startCol = 20)
writeData(wb, sheet = "2022", 
          FUND_GAM_Beta_22_DF,
          startRow = 13, startCol = 20)
writeData(wb, sheet = "2022", 
          FUND_GAM_Treynor_22_DF,
          startRow = 16, startCol = 20)
writeData(wb, sheet = "2022", 
          FUND_GAM_Alpha_22_DF,
          startRow = 19, startCol = 20)
writeData(wb, sheet = "2022", 
          FUND_GAM_SRRI_22_DF,
          startRow = 22, startCol = 20)

setColWidths(wb, "2022", cols = 1:30, widths = "auto")


# 2023
addWorksheet(wb, "2023")

# S&P500
writeData(wb, sheet = "2023", 
          "BENCHMARK",
          startRow = 2, startCol = 2)
writeData(wb, sheet = "2023", 
          SP500_Ret_23_DF,
          startRow = 4, startCol = 2)
writeData(wb, sheet = "2023", 
          SP500_SD_23_DF,
          startRow = 7, startCol = 2)
writeData(wb, sheet = "2023", 
          SP500_Sr_23_DF,
          startRow = 10, startCol = 2)

# PTF_MINVAR
writeData(wb, sheet = "2023", 
          "PTF MIN VAR",
          startRow = 2, startCol = 7)
writeData(wb, sheet = "2023", 
          weights_minvar_23_DF,
          startRow = 4, startCol = 5,
          rowNames = TRUE)
writeData(wb, sheet = "2023", 
          PTF_MINVAR_Ret_23_DF,
          startRow = 4, startCol = 8)
writeData(wb, sheet = "2023", 
          PTF_MINVAR_SD_23_DF,
          startRow = 7, startCol = 8)
writeData(wb, sheet = "2023", 
          PTF_MINVAR_Sr_23_DF,
          startRow = 10, startCol = 8)

# PTF_MAXSR
writeData(wb, sheet = "2023", 
          "PTF MAX SR",
          startRow = 2, startCol = 13)
writeData(wb, sheet = "2023", 
          weights_maxsr_23_DF,
          startRow = 4, startCol = 11,
          rowNames = TRUE)
writeData(wb, sheet = "2023", 
          PTF_MAXSR_Ret_23_DF,
          startRow = 4, startCol = 14)
writeData(wb, sheet = "2023", 
          PTF_MAXSR_SD_23_DF,
          startRow = 7, startCol = 14)
writeData(wb, sheet = "2023", 
          PTF_MAXSR_Sr_23_DF,
          startRow = 10, startCol = 14)

# FUND_GAM
writeData(wb, sheet = "2023", 
          "FOND GAM",
          startRow = 2, startCol = 19)
writeData(wb, sheet = "2023", 
          weights_ew_23_DF,
          startRow = 4, startCol = 17,
          rowNames = TRUE)
writeData(wb, sheet = "2023", 
          FUND_GAM_Ret_23_DF,
          startRow = 4, startCol = 20)
writeData(wb, sheet = "2023", 
          FUND_GAM_SD_23_DF,
          startRow = 7, startCol = 20)
writeData(wb, sheet = "2023", 
          FUND_GAM_Sr_23_DF,
          startRow = 10, startCol = 20)
writeData(wb, sheet = "2023", 
          FUND_GAM_Beta_23_DF,
          startRow = 13, startCol = 20)
writeData(wb, sheet = "2023", 
          FUND_GAM_Treynor_23_DF,
          startRow = 16, startCol = 20)
writeData(wb, sheet = "2023", 
          FUND_GAM_Alpha_23_DF,
          startRow = 19, startCol = 20)
writeData(wb, sheet = "2023", 
          FUND_GAM_SRRI_23_DF,
          startRow = 22, startCol = 20)

setColWidths(wb, "2023", cols = 1:30, widths = "auto")


#path = paste0("C:/Users/quent/OneDrive/Documents/Etudes/M2/Cours/R/DM/", 
#              "Gustave_AM_reporting.xlsx")
path = paste0("C:/Users/quent/OneDrive/Documents/Etudes/M2 MQF/Classes/Track MQF/Intro to DS with R/Gustave_AM_reporting.xlsx")

# C:\Users\quent\OneDrive\Documents\Etudes\M2 MQF\Classes\Track MQF\Intro to DS with R

saveWorkbook(wb, path, overwrite = TRUE)

# END STEP 4 -------------------------------------------------------------------



