import sys
import pandas as pd
import yfinance as yf

def fetch_yahoo_monthly(tickers, years=7):
    dfs = []
    for sym in tickers:
        d = yf.download(sym, period=f"{years}y", interval="1mo", auto_adjust=False, progress=False)
        if d is None or d.empty:
            print(f"Warning: no data for {sym}")
            continue
        col = "Adj Close" if "Adj Close" in d.columns else "Close"
        s = d[[col]].rename(columns={col: sym})
        s.index = s.index.to_period("M").to_timestamp("M")
        dfs.append(s)
    prices = pd.concat(dfs, axis=1).sort_index()
    return prices

if __name__ == "__main__":
    tickers = ["AAPL","MSFT","JPM","KO","MCD","XOM","CAT","JNJ","WMT","DIS","^GSPC"]
    years = int(sys.argv[1]) if len(sys.argv) > 1 else 7
    prices = fetch_yahoo_monthly(tickers, years=years)
    prices.to_csv("prices_monthly.csv", float_format="%.6f")
    print("Saved prices_monthly.csv with month-end index.")
