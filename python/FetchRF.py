from pandas_datareader import data as pdr
import pandas as pd
from datetime import datetime

def fetch_fred_dtb3_monthly(start=None, end=None):
    if start is None: 
        start = datetime(2000, 1, 1)
    if end is None: 
        end = datetime.today()

    rf_daily = pdr.DataReader("DTB3", "fred", start, end)
    rf_m = rf_daily.resample("M").last()
    rf_m.index = rf_m.index.to_period("M").to_timestamp("M")
    rf_m["RF"] = rf_m["DTB3"].astype(float)/100.0/12.0  # convert annual % to monthly decimal
    return rf_m[["RF"]]

if __name__ == "__main__":
    rf = fetch_fred_dtb3_monthly()
    rf.to_csv("rf_dtb3_monthly.csv", float_format="%.6f")
    print("Saved rf_dtb3_monthly.csv (monthly decimals, annual/12).")
