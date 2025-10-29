import io
import zipfile
import requests
import pandas as pd

def fetch_ff_monthly():
    url = "https://mba.tuck.dartmouth.edu/pages/faculty/ken.french/ftp/F-F_Research_Data_Factors_CSV.zip"
    r = requests.get(url, timeout=30)
    z = zipfile.ZipFile(io.BytesIO(r.content))
    csv = [n for n in z.namelist() if n.endswith(".csv")][0]

    ff = pd.read_csv(z.open(csv), skiprows=3)
    ff = ff.rename(columns={ff.columns[0]: "Date"})

    # Remove any annual summary rows
    cut = ff[ff["Date"].astype(str).str.contains("Annual", na=False)].index
    if len(cut):
        ff = ff.loc[:cut[0]-1]

    # Convert dates
    def parse(x):
        x = str(x).strip()
        if len(x) == 6 and x.isdigit():
            y, m = int(x[:4]), int(x[4:])
            return pd.Timestamp(y, m, 1) + pd.offsets.MonthEnd(0)
        return pd.NaT

    ff["Date"] = ff["Date"].apply(parse)
    ff = ff.dropna(subset=["Date"]).set_index("Date").sort_index()

    # Convert to decimal
    for c in ["Mkt-RF", "SMB", "HML", "RF"]:
        ff[c] = ff[c].astype(float)/100.0

    return ff[["Mkt-RF", "SMB", "HML", "RF"]]

if __name__ == "__main__":
    ff = fetch_ff_monthly()
    ff.to_csv("ff_factors_monthly.csv", float_format="%.6f")
    print("Saved ff_factors_monthly.csv (monthly decimals).")
