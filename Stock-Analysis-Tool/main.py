import pandas as pd
import yfinance as yf
import math
import time
import numpy as np
import xlsxwriter
from statistics import mean
from scipy import stats
import matplotlib.pyplot as plt
import seaborn as sns
from scipy.stats import norm
from scipy.stats import spearmanr
from scipy.stats import trim_mean
from sklearn.linear_model import LinearRegression


def main():

    # Wczytywaniapliku CSV
    stocks = pd.read_csv("sp_500_stocks.csv")

    # Grupowanie tickerów po 100
    symbol_groups = list(chunks(stocks["Ticker"], 100))
    symbol_string = [" ".join(x) for x in symbol_groups]
    
    # Tworzymy pusty DataFrame
    my_columns = [
        "Ticker", 
        "Prefered to Buy", 
        "Price", 
        "Prefered to Sell", 
        "Number of shares to Buy", 
        "Price-to-Earning Ratio", 
        "PE Percentile", 
        "Price-to-Book Ratio", 
        "PB Percentile", 
        "Price-to-Sales Ratio", 
        "PS Percentile",
        "EV/EBITDA", 
        "EV/EBITDA Percentile", 
        "EV/GP", 
        "EV/GP Percentile", 
        "RV Score"
        ]

    final_dataframe = pd.DataFrame(columns = my_columns)

    # Pobieranie danych dla każdego z tickerów
    for symbols in symbol_string:
        print(f"\nPobieranie danych dla grupy: {symbols}")

        # Pobieranie danch giełdowych z yfinance
        try:
            time.sleep(10)

            # Pętla przetwarzająca tickery
            for ticker in symbols.split(" "):
                c = yf.Ticker(ticker)
                info = c.info
                
                # Price
                price = info.get("previousClose", None)

                # Price-to-Earning Ratio (P/E)
                pe_raw = info.get("trailingPE", None)
                pe_ratio = round(pe_raw, 2) if isinstance(pe_raw, (int, float)) else None

                # Price-to-Book Ratio (P/B) 
                pb_raw = info.get("priceToBook", None)
                pb_ratio = round(pb_raw, 2) if isinstance(pb_raw, (int, float)) else None

                # Price-to-Sales Ratio (P/S)
                ps_raw = info.get("priceToSalesTrailing12Months", None)
                ps_ratio = round(ps_raw, 2) if isinstance(ps_raw, (int, float)) else None

                # EV/EBITDA
                ee_raw = info.get("enterpriseToEbitda", None)
                ee_ratio = round(ee_raw, 2) if isinstance(ee_raw, (int, float)) else None

                # EV/GP
                ev_raw = info.get("enterpriseValue", None)
                ev = round(ev_raw, 2) if isinstance(ev_raw, (int, float)) else None
                gp_raw = info.get("grossProfits", None)
                gp = round(gp_raw, 2) if isinstance(gp_raw, (int, float)) else None
                ev_gp = round(ev / gp, 2)

                # Dopisanie do DataFrame
                new_row = pd.DataFrame([[
                                        ticker, 
                                        "N/A", price, "N/A", "N/A", 
                                        pe_ratio, "N/A", 
                                        pb_ratio, "N/A", 
                                        ps_ratio, "N/A", 
                                        ee_ratio, "N/A", 
                                        ev_gp, "N/A", 
                                        "N/A"
                                        ]], columns = my_columns)
                final_dataframe = pd.concat([final_dataframe, new_row], ignore_index = True)

        # Ponowienie w przypadku błędu
        except Exception as e:
            print(f"Błąd pobierania danych dla {symbols}: {e}")
            time.sleep(30)
            continue

    # Obliczanie procentów    
    metrics = {         
        "Price-to-Earning Ratio" : "PE Percentile", 
        "Price-to-Book Ratio" : "PB Percentile", 
        "Price-to-Sales Ratio" : "PS Percentile",
        "EV/EBITDA" : "EV/EBITDA Percentile", 
        "EV/GP" : "EV/GP Percentile"
    }

    for metric in metrics.keys():
        for row in final_dataframe.index:
            value = final_dataframe.loc[row, metric]
            if isinstance(value, (int, float)):
                percentile = stats.percentileofscore(final_dataframe[metric].dropna(), value)
                final_dataframe.loc[row, metrics[metric]] = round(percentile, 2)
            else:
                final_dataframe.loc[row, metrics[metric]] = "N/A" 

    print(final_dataframe)
    
    # Obliczenie RV Score
    for row in final_dataframe.index:

        rv_score = []
        for metric in metrics.keys():
            rv_score.append(final_dataframe.loc[row, metrics[metric]])
                            
        final_dataframe.loc[row, "RV Score"] = mean(rv_score)

    # Zostawienie 50 spółek które najlepeij spełniają nasze wymogi
    columns_to_check = [
        "Price-to-Earning Ratio", "PE Percentile", 
        "Price-to-Book Ratio", "PB Percentile", 
        "Price-to-Sales Ratio", "PS Percentile",
        "EV/EBITDA", "EV/EBITDA Percentile", 
        "EV/GP", "EV/GP Percentile", 
        "RV Score"
    ]

    # Odrzucamy wiersze z jakąkolwiek wartością mniejszą od 0
    final_dataframe = final_dataframe[
        final_dataframe[columns_to_check].apply(
            lambda row: all(isinstance(val, (int, float)) and val >= 0 for val in row), axis=1
        )
    ]
    
    final_dataframe = final_dataframe.sort_values("RV Score")
    final_dataframe = final_dataframe[:50]
    final_dataframe.reset_index(inplace = True)
    final_dataframe.drop("index", axis=1, inplace = True)
        
    # Podanie wartości portfela inwestycyjnego
    portfolio_size = portfolio_input()
    position_size = float(portfolio_size)/len(final_dataframe.index)
    for i in final_dataframe.index:
        final_dataframe.loc[i, "Number of shares to Buy"] = math.floor(position_size / final_dataframe.loc[i, "Price"])

    # 50 spółek dodanie do listy
    sorted_list = list(final_dataframe["Ticker"])
    print(sorted_list)

    # Dodanie "Prefered to Buy" i "Prefered to Sell"
    for ticker in sorted_list:
        df = yf.download(ticker, period="5y", interval="1d")

        # Resampling do tygodniowych danych
        weekly_data = df.resample("W").agg({("Close", ticker): ["first", "last"]})
        weekly_data.columns = ["start_of_week_price", "end_of_week_price"]
        weekly_data["start_of_the_week"] = weekly_data.index - pd.offsets.Week(weekday=0) + pd.offsets.BDay(0)
        weekly_data["end_of_the_week"] = weekly_data.index - pd.offsets.Week(weekday=4) + pd.offsets.BDay(0)
        
        # Obliczanie zmian
        weekly_data["PriceChange"] = weekly_data["start_of_week_price"] - weekly_data["end_of_week_price"]
        weekly_data["PercentChange"] = (weekly_data["PriceChange"] / weekly_data["start_of_week_price"]) * 100
        percent_changes = weekly_data["PercentChange"].dropna()

        
        # Przycięcie danych (10% z każdej strony)
        sorted_changes = percent_changes.sort_values()
        cut_len = int(0.1 * len(sorted_changes))
        trimmed_changes = sorted_changes.iloc[cut_len : len(sorted_changes) - cut_len]

        # Dopasowanie rozkładu tylko do przyciętych danych
        mu, sigma = norm.fit(trimmed_changes)
        
        # Obliczenia sygnałów kupna/sprzedaży
        mu_plus_1s = mu + 2 * sigma
        mu_minus_1s = mu - 2 * sigma
        base_price = weekly_data["start_of_week_price"].mean()
        
        buy_price = base_price * (1 + mu_minus_1s / 100)
        sell_price = base_price * (1 + mu_plus_1s / 100)

        # Dodanie do df
        final_dataframe.loc[final_dataframe["Ticker"] == ticker, "Prefered to Buy"] = round(buy_price, 2)
        final_dataframe.loc[final_dataframe["Ticker"] == ticker, "Prefered to Sell"] = round(sell_price, 2)
    
    print(final_dataframe)

    # Zapis do Excela
    with pd.ExcelWriter("sp500_value_report.xlsx", engine="xlsxwriter") as writer:
        final_dataframe.to_excel(writer, sheet_name="RV Strategy", index=False)

        workbook = writer.book
        worksheet = writer.sheets["RV Strategy"]

        # Formatowanie
        money_format = workbook.add_format({'num_format': '$0.00'})
        integer_format = workbook.add_format({'num_format': '0'})
        float_format = workbook.add_format({'num_format': '0.00'})
        text_format = workbook.add_format({'num_format': '@'})

        worksheet.set_column('A:A', 10, text_format)
        worksheet.set_column('B:B', 12, money_format)
        worksheet.set_column('C:C', 12, money_format)
        worksheet.set_column('D:D', 12, money_format)
        worksheet.set_column('E:E', 22, integer_format)
        worksheet.set_column('F:F', 22, float_format)
        worksheet.set_column('G:G', 15, float_format)
        worksheet.set_column('H:H', 22, float_format)
        worksheet.set_column('I:I', 15, float_format)
        worksheet.set_column('J:J', 22, float_format)
        worksheet.set_column('K:K', 15, float_format)
        worksheet.set_column('L:L', 22, float_format)
        worksheet.set_column('M:M', 15, float_format)
        worksheet.set_column('N:N', 22, float_format)
        worksheet.set_column('O:O', 15, float_format)
        worksheet.set_column('P:P', 10, float_format)      

# Funkcja do dzielenia listy na grupy po "n" elementów
def chunks(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i + n]

# Funkcja do wprowadzania wartości portfela inwestycyjnego
def portfolio_input():
    global portfolio_size
    while True:
        portfolio_size = input("Podaj wartość portfela inwestycyjnego: ")
        try:
            float(portfolio_size)
            break
        except ValueError:
            print("Podaj wartość portfela w cyfrach.")
    print(f"Wartość portfela inwestycyjnego wynośi: {portfolio_size}")
    return(portfolio_size)

main()