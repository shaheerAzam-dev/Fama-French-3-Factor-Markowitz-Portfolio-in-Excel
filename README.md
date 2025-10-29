Fama–French 3-Factor + Markowitz Portfolio in Excel
===================================================

Independent Project: Portfolio Construction & Factor Modeling

Introduction
------------

This project demonstrates the full workflow of portfolio construction for 10 large NYSE stocks, combining Fama–French 3-factor regression analysis with Markowitz mean-variance optimization in Excel.

We leverage historical stock and factor data to:

*   Estimate expected returns using factor models
*   Calculate portfolio risk via covariance matrices
*   Construct optimal portfolios (Minimum Variance and Tangency) using Excel Solver
*   Visualize results and provide a clear, report-ready summary
    
The goal is to connect academic theory with practical, hands-on portfolio management in a transparent, reproducible way.

### Quick Links

01. Raw Data: Yahoo Finance (Adjusted Close) — [https://finance.yahoo.com/](https://finance.yahoo.com/)
02. Fama–French Factors: Kenneth French Data Library — [https://mba.tuck.dartmouth.edu/pages/faculty/ken.french/data\_library.html](https://mba.tuck.dartmouth.edu/pages/faculty/ken.french/data_library.html)
03. Risk-Free Rate: FRED DTB3 (3-month T-bill) — [https://fred.stlouisfed.org/series/DTB3](https://fred.stlouisfed.org/series/DTB3)
    

Background
----------

### Project Scope

*   Assets: AAPL, MSFT, JPM, KO, MCD, XOM, CAT, JNJ, WMT, DIS
*   Market Proxy: S&P 500 (^GSPC)
*   Timeframe: Last 7 years, monthly frequency

This setup ensures sufficient diversification, yet keeps analysis manageable and visually clear in Excel.

*   Why 10 stocks: Beyond 10–15 reasonably uncorrelated stocks, incremental benefit from diversification is minimal.
*   Why 7 years: Long enough to estimate covariances and betas without excessive noise, recent enough to reflect current market conditions.
*   Why Excel: Transparent formulas, visible regressions, and hands-on Solver optimization demonstrate all steps clearly.
  

Ask
---

### Business Task

Construct and analyze an optimal portfolio of 10 large-cap NYSE stocks using factor models and mean-variance optimization.

### Analysis Questions

1.  How do the Fama–French 3-factor model estimates compare to historical returns?
2.  How does covariance between assets affect portfolio risk?
3.  What are the optimal portfolio weights for minimum variance and maximum Sharpe ratio (tangency portfolio)?
4.  How do visualizations (efficient frontier, asset scatter, MVP/TP pies) help interpret results?
    

Prepare
-------

### Data Sources

*   Stock Prices: Yahoo Finance (monthly adjusted close)
*   Factors: Fama–French 3-factor monthly data
*   Risk-Free Rate: FRED 3-month T-bill (annual/12 for monthly decimals)
    

### Data Organization (Excel Sheets)

*   RawPrices – stock + S&P 500 prices
*   Factors\_RF – MKT-RF, SMB, HML, RF
*   Returns – arithmetic & log returns, market excess returns
*   ExcessReturns – stock excess returns (R − RF)
*   Regressions – Excel ToolPak regression outputs & residuals
*   ExpectedReturns – monthly & annualized E\[R\]
*   CovMatrix – 10×10 covariance of excess returns    
*   PortfolioCalc – weights, expected return, variance, stdev, Sharpe    
*   Solver\_MVP / Solver\_TP – Solver instructions & reference results
*   Frontier – efficient frontier table for plotting    
*   Charts – frontier chart, asset scatter, MVP/TP pies

<img width="1140" height="48" alt="image" src="https://github.com/user-attachments/assets/56555d4c-85fc-4ab1-9906-9556d916772a" />
    



### Step 1 – Data Pull
----------------------

This step gathers all the raw data that powers the portfolio model.  
To make the workflow efficient and reproducible, I used **three small Python scripts** — one for each data source — to pull, clean, and save monthly CSV files.  

These scripts do not perform any analysis; they simply automate the collection and alignment of inputs that are later used in Excel.


#### 🔹 Scripts Overview

| Script | Purpose | Output File | Source |
|--------|----------|--------------|--------|
| `[fetch_prices.py]()` | Downloads monthly **adjusted close prices** for 10 stocks and the **S&P 500 (^GSPC)** | `[prices_monthly.csv]()` | Yahoo Finance |
| `[fetch_factors.py]()` | Retrieves **Fama–French 3-Factor data** (MKT−RF, SMB, HML, RF) | `[ff_factors_monthly.csv]()` | Kenneth R. French Data Library |
| `[fetch_rf.py]()` | Pulls **3-month Treasury Bill (DTB3)** data, converts it to **monthly decimal RF** | `[rf_dtb3_monthly.csv]()` | FRED (Federal Reserve) |

Each script automatically sets the date range (last 7 years) and outputs a clean CSV ready for import into Excel.  


#### ⚙️ Workflow  

1. **Run the scripts**  
   Execute each `.py` file from your terminal or IDE:  
   ```bash
   python fetch_prices.py
   python fetch_factors.py
   python fetch_rf.py 
  Each script saves a CSV file in your working directory.

2.  **Import CSVs into Excel**
    
    *   prices\_monthly.csv → paste into the **RawPrices** sheet
    *   ff\_factors\_monthly.csv → paste into the **Factors\_RF** sheet        
    *   rf\_dtb3\_monthly.csv → cross-check that the monthly RF values line up with the factor data
        
3.  **Align on Month-End Dates**
    
    *   Use **month-end alignment** across all sheets (RawPrices, Factors\_RF, Returns).        
    *   Drop any months that don’t appear in all three datasets (inner join).        
    *   This ensures that every regression and covariance calculation uses perfectly matched rows.

<img width="1213" height="351" alt="image" src="https://github.com/user-attachments/assets/21b9d117-ff08-4383-b298-324e163586ae" />
<img width="506" height="351" alt="image" src="https://github.com/user-attachments/assets/65406555-ded8-4678-8301-4e6fa030468a" />


### Step 2 – Return Calculations
--------------------------------

This step converts prices and macro inputs into the return series that power both the factor regressions and the optimizer. Accuracy and alignment here are critical—every downstream estimate depends on these returns.

**What we compute and why:**
- Arithmetic returns: used for factor regressions and portfolio optimization because they aggregate linearly across assets.
- Excess returns: subtract the contemporaneous monthly risk-free rate (RF) so that returns are measured relative to cash—this becomes the dependent variable for the Fama–French regressions.
- Market excess returns: same idea for the market proxy (^GSPC), used as an explanatory factor.

**Returns Formula**  
<img width="213" height="36" alt="image" src="https://github.com/user-attachments/assets/f6902a12-314e-440f-93a1-848fd5ac7713" />  
<img width="1319" height="351" alt="image" src="https://github.com/user-attachments/assets/e4398a94-45ed-4d1c-a453-7daae75887a8" />

**Excess Returns Formula**  
<img width="175" height="36" alt="image" src="https://github.com/user-attachments/assets/358f907f-f2e5-4ad0-b60e-d884991f81dd" />  
<img width="1239" height="351" alt="image" src="https://github.com/user-attachments/assets/3e369e59-e306-4d74-a583-3a90938bd1b7" />

**Implementation notes:**
- Apply arithmetic return formulas to all stocks and to the market proxy (^GSPC).
- Align dates across RawPrices, Factors_RF, and Returns; drop months that don’t appear in all three datasets (inner join).
- Ensure the risk-free rate is in monthly decimal terms (e.g., 0.002 rather than 0.2%).
- Confirm dividends are included via Adjusted Close; this ensures total-return consistency across assets.

**Quality checks:**
- Plot quick histograms of monthly returns to spot anomalies or fat tails.
- Verify that Average(Excess Return) = Average(Return) − Average(RF) for multiple tickers.
- Remove any rows with missing or malformed values; #DIV/0! or blanks will break regressions and MMULT.

**Deliverables from this step:**
- Returns sheet with monthly arithmetic returns for all assets and the market.
- ExcessReturns sheet with monthly excess returns for all assets and the market.


### Step 3 – Fama–French 3-Factor Regressions
---------------------------------------------

The goal here is to estimate how each stock’s excess return loads on the Fama–French factors. We run one regression per stock, using monthly data across the aligned sample window.

For each stock i, regress monthly excess returns on the Fama–French factors:

$$
(R_i - R_F) = \alpha + \beta_{MKT}(MKT - R_F) + \beta_{SMB}(SMB) + \beta_{HML}(HML) + \varepsilon
$$

Procedure in Excel (ToolPak):
1. Data → Data Analysis → Regression    
2. Y Range: Stock_i_Excess_Returns    
3. X Range: columns for MKT−RF, SMB, HML    
4. Check “Labels” if headers included; include intercept    
5. Output residuals and summary statistics

<img width="1131" height="246" alt="image" src="https://github.com/user-attachments/assets/92e3d7fa-a206-4fcb-a41e-5c112670895b" />

Record for each stock:
- α, β_MKT, β_SMB, β_HML    
- R², standard error
- Residuals (idiosyncratic component)

<img width="607" height="386" alt="image" src="https://github.com/user-attachments/assets/9e7f5549-38a5-4e7d-9638-4b833a863400" />

Interpretation:
- β_MKT: sensitivity to broad market risk.
- β_SMB: tilt toward small-cap vs. large-cap characteristics.
- β_HML: tilt toward value vs. growth characteristics.
- α: historical intercept; for forward-looking E[R], set α = 0 to avoid in-sample overfitting.

Good practice and diagnostics:
- Use the exact same date range across all regressions; mismatches distort betas.
- Inspect residuals for outliers or patterns that might indicate missing data or alignment issues.
- ToolPak provides OLS with classical SEs; for research-grade inference consider robust SEs (e.g., Newey–West) in a stats package.
- Remember: high R² indicates explanatory power of variation, not necessarily higher expected return.

Deliverables from this step:
- Regressions sheet with full ToolPak outputs (one per stock).
- A concise Betas summary table (α, βs, R²) covering all 10 stocks.
- Residuals saved for optional idiosyncratic risk diagnostics.


### Step 4 – Expected Returns
-----------------------------

We translate factor exposures (betas) into forward-looking expected returns by combining them with estimated factor premia (historical means over the same 7-year window). This keeps regression and premia estimation consistent.

Compute expected monthly returns from the factor model by multiplying estimated betas with the historical factor premia:

Calculated from Factors_RF  
<img width="438" height="71" alt="image" src="https://github.com/user-attachments/assets/e3094c8c-544f-4eec-8c22-b13c8f877ddc" />

Monthly Return formula:

$$
E[R_i] = R_F + \beta_{MKT} \, E(MKT - RF) + \beta_{SMB} \, E(SMB) + \beta_{HML} \, E(HML)
$$

Geometric Annualization:

$$ (1 + E_{monthly})^{12} - 1 $$

Linear (approximation) Annualization:

$$ 12 \times E_{monthly} $$

Notes:
- RF should be the monthly average over the same sample used for factor means.
- Setting α = 0 is a conservative and standard forward-looking assumption.
- Cross-check: compare FF3-implied E[R_i] with each stock’s historical average excess return to understand model vs. realized return gaps.

Quality checks:
- Confirm factor means were computed on the identical rows used in regressions.
- Sanity-check rankings: with a positive market premium, higher β_MKT should typically raise E[R_i].
- Store both monthly and annualized E[R_i] for downstream plotting and reporting.

Deliverables from this step:
- ExpectedReturns sheet with monthly E[R_i], and geometric/linear annualized values for all 10 stocks.


### Step 5 – Covariance Matrix
------------------------------

The covariance matrix Σ captures how assets co-move and is the backbone of portfolio risk. Diversification works when off-diagonal covariances are low or negative, reducing overall variance.

Build a 10×10 covariance matrix (Σ) using monthly excess returns:
- Diagonal entries: VAR.P of each stock’s excess returns    
- Off-diagonal entries: COV.P of each stock pair’s excess returns    
- Keep Σ in monthly units for optimization; annualize only for charts

<img width="515" height="36" alt="image" src="https://github.com/user-attachments/assets/01f77c98-e6d8-4e26-9b39-303a501bf713" />  
<img width="1112" height="386" alt="image" src="https://github.com/user-attachments/assets/5ab51a36-684a-4843-b012-ab2b9d557885" />

Tips:
- Use the exact same aligned ExcessReturns range used in Step 3—consistency is crucial.
- Prefer population formulas (VAR.P, COVARIANCE.P) for stability in smaller samples.
- Optional: Apply simple shrinkage toward the diagonal (e.g., 10–20%) if Σ appears unstable.
- Verify symmetry: Σ must be symmetric; if not, re-check your cell references.

Quality checks:
- Build a correlation heatmap from the same range; verify signs/magnitudes are sensible.
- Confirm no NA or text cells in the Σ range; they will break MMULT.

Deliverables from this step:
- CovMatrix sheet containing a clean, symmetric monthly covariance matrix Σ for all 10 assets.


### Step 6 – Portfolio Optimization
-----------------------------------

With E[R] from Step 4 and Σ from Step 5, we compute any portfolio’s expected return and variance, then use Solver to find optimal weights under constraints.

Set up weight base (initially with equal weights) and Expected Returns table (from FF3)  
<img width="348" height="421" alt="image" src="https://github.com/user-attachments/assets/ee308e79-d6c8-4862-9c61-e7b80dd72fc3" />  <img width="289" height="386" alt="image" src="https://github.com/user-attachments/assets/81abaee3-4640-4d21-883e-78581756d908" />

Set up formulas:
- Portfolio return (monthly):  $= \text{SUMPRODUCT}(weights\_range, expected\_returns\_range)$  
- Portfolio variance:  $= \text{MMULT}(\text{TRANSPOSE}(weights\_range), \text{MMULT}(\Sigma\_range, weights\_range))$  
- Portfolio standard deviation:  $= \text{SQRT}(portfolio\_variance\_cell)$  
- Sharpe ratio (monthly):  $= \dfrac{PortfolioReturn - RF\_cell}{PortfolioStdev}$

<img width="737" height="246" alt="image" src="https://github.com/user-attachments/assets/2d86832a-6e5e-485f-a74a-b4a30b51cb67" />

Constraints (long-only base case):
- SUM(weights) = 1    
- Each weight ≥ 0

Solver settings:
- MVP (Minimum Variance Portfolio): Set objective to minimize variance  
<img width="562" height="552" alt="image" src="https://github.com/user-attachments/assets/4ae9642c-468f-4a26-a0b1-63d8628f4be5" />

- TP (Tangency Portfolio): Set objective to maximize Sharpe  
<img width="562" height="550" alt="image" src="https://github.com/user-attachments/assets/113ae250-425f-42f0-b977-934f33cd7854" />

- Engine: GRG Nonlinear    
- Starting values: equal weights or MVP solution

Operational tips:
- For Tangency, ensure RF_cell references the monthly risk-free rate used throughout.
- If Solver struggles to converge, try:
  - Using MVP solution as a warm start,
  - Applying per-asset caps (e.g., ≤ 40–50%) to regularize,
  - Switching to the Evolutionary engine as a fallback.
- After solving, round weights only for presentation; keep full-precision weights for calculations.

Validation:
- Confirm Solver converges without binding inconsistencies.
- Inspect weight concentration; long-only TP often clusters in a few high E[R]/σ names.
- Check MVP variance is below most single-name variances; if not, re-check Σ and date alignment.

Deliverables from this step:
- PortfolioCalc sheet with dynamic portfolio metrics.
- Solver_MVP and Solver_TP setups documented with final weights and performance stats.


### Step 7 – Efficient Frontier & Visuals
-----------------------------------------

This step translates the optimizer into a full picture of risk–return trade-offs. You’ll generate a curve of feasible portfolios by targeting return levels and solving for minimum variance at each target.

Efficient Frontier:
- Create a column of target monthly returns spanning MVP to max observed E[R]    
- For each target, use Solver to minimize variance subject to:
  - SUM(weights)=1
  - weights≥0
  - SUMPRODUCT(weights,E[R])=Target
- Record portfolio stdev and annualize for plotting

<img width="561" height="351" alt="image" src="https://github.com/user-attachments/assets/c6b298a3-acca-47ee-9696-933ce0fc4f89" />

Charts to include:
- Efficient Frontier (annualized)  
<img width="1547" height="938" alt="image" src="https://github.com/user-attachments/assets/49cb6622-7e5f-4b87-a152-ac19cd7ebec6" />

- Asset scatter (annualized return vs. stdev)  
<img width="1219" height="936" alt="image" src="https://github.com/user-attachments/assets/e72bc69a-b774-4973-9f90-b7c9e3052557" />

- Pie charts for MVP and TP  
<img width="1536" height="675" alt="image" src="https://github.com/user-attachments/assets/4042f24e-aefb-43e3-ba73-5fe8e5526c9b" />

- FF Betas by Stock  
<img width="2226" height="609" alt="image" src="https://github.com/user-attachments/assets/636b5a32-7f28-46ca-8b45-39f7938d1ab6" />

How to read the visuals:
- The asset scatter shows each stock’s standalone position; the frontier traces the best achievable trade-off via diversification.
- MVP lies at the far-left of the frontier (lowest variance). Tangency (max Sharpe) is the point where a straight line from RF is just tangent to the frontier.
- Pie charts reveal portfolio intuition—defensive names dominate MVP; higher-return exposures tilt TP.

Quality checks:
- Ensure every frontier point uses identical constraints and correct references.
- Verify monotonicity: annualized stdev should generally rise with higher target returns.

Deliverables from this step:
- Frontier sheet with a clean table of target returns, monthly σ, annualized stats, and solved weights.
- Charts sheet containing the finalized frontier, asset scatter, and MVP/TP pies ready for inclusion in the report.





