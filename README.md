# S&P 500 Quantitative Value Analysis Tool 📈

## Overview
This is a Python-based algorithmic trading tool designed to automate the search for undervalued stocks within the S&P 500 index. It combines **Fundamental Analysis** (Value Investing) with **Technical Analysis** (Volatility) to generate a comprehensive investment report.

The script calculates a "Relative Value (RV) Score" for each stock and determines statistically optimal entry and exit points.

## Key Features

* **Automated Data Collection:** Scrapes real-time financial data for S&P 500 companies using `yfinance`.
* **Composite RV Score:** Ranks stocks based on an algorithmic score derived from key valuation metrics:
    * Price-to-Earnings (P/E)
    * Price-to-Book (P/B)
    * Price-to-Sales (P/S)
    * EV/EBITDA
    * EV/Gross Profit
* **Statistical Trade Signals:** Calculates "Preferred Buy" and "Preferred Sell" prices using Normal Distribution (2 Standard Deviations) on historical weekly price changes.
* **Portfolio Management:** Calculates the exact number of shares to buy based on the user's total portfolio size.
* **Excel Reporting:** Exports a fully formatted `.xlsx` dashboard using `xlsxwriter`.

## Technologies
* **Python 3.x**
* **Pandas & NumPy** (Data Manipulation)
* **SciPy** (Statistical Analysis)
* **YFinance** (Market Data API)
* **XlsxWriter** (Excel Automation)

## How to Run

1.  **Clone the repository:**
    ```bash
    git clone [https://github.com/TWOJ-NICK/TWOJE-REPO.git](https://github.com/TWOJ-NICK/TWOJE-REPO.git)
    ```
2.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```
3.  **Check the input file:**
    Ensure that `sp_500_stocks.csv` is present in the main directory.
4.  **Run the script:**
    ```bash
    python main.py
    ```
5.  **Follow the prompt:**
    Enter your total portfolio value (e.g., `100000`) when asked. The script will generate `sp500_value_report.xlsx`.

## Strategy Logic
1.  **Filtering:** The script selects the top 50 stocks with the lowest (best) RV Score.
2.  **Valuation:** It compares each stock against the broader market using percentile ranking.
3.  **Volatility Check:** By analyzing 5 years of weekly price data, the script identifies price levels that are statistically significant deviations from the mean (±2σ), serving as potential buy/sell targets.

---
*Disclaimer: This tool is for educational purposes only and does not constitute financial advice.*
