# hk-equity-excel-toolkit
hk equity analysis
HK Equity Analytics Toolkit
Excel + VBA portfolio project  |  Data: Yahoo Finance daily adjusted close, 2 Oct 2025 - 16 Oct 2026 (256 trading days)

What it does
1. Dashboard: pick any of 5 HK blue chips (or the HSI) from a dropdown; all KPIs, charts and an MA-crossover backtest update live.
2. Stats: risk/return table for every ticker - annualised return & volatility, Sharpe, beta vs HSI, max drawdown, 52-week range, ranking.
3. Correlation: 6x6 correlation matrix of daily returns built with CORREL + INDEX/MATCH, colour-scaled.
4. Monthly: formula-driven pivot - monthly returns, avg daily volume (AVERAGEIFS), up-days (COUNTIFS) by ticker x month.
5. Data: long-format Excel Table (tblData) - the source for a native PivotTable (see "Build the PivotTable" below).
6. VBA (StockToolkit.bas): UDFs MaxDrawdown() / AnnVol(), a macro that loops every ticker and writes a report sheet, PDF export, CSV import.

Colour convention
Blue text = hard-coded input  |  Black = formula  |  Green = cross-sheet link  |  Yellow fill = user input cell

How to verify
- Every number outside Prices/Data is a live formula (no pasted values). Click any cell to audit; Formulas > Trace Precedents.
- Spot-check prices: Yahoo Finance > ticker > Historical Data, e.g. 0700.HK close on 16 Oct 2026 = 417.80.
- Change Dashboard!C4 (ticker) or C6/C7 (MA windows): the chart, signal and backtest recompute instantly.

Key functions used
INDEX/MATCH, OFFSET, SUMPRODUCT, AVERAGEIFS, COUNTIFS, SLOPE, CORREL, STDEV, EOMONTH, RANK, IFERROR, TEXT, data validation, conditional formatting, Excel Table, charts



