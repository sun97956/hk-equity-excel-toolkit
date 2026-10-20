Attribute VB_Name = "StockToolkit"
'==============================================================================
' StockToolkit  -  VBA module for HK_Equity_Analytics_Toolkit
' Import: Alt+F11 > File > Import File > StockToolkit.bas, then save as .xlsm
'
' Contents
'   UDFs (use in any cell):
'     =MaxDrawdown(Prices!B2:B257)        worst peak-to-trough decline
'     =AnnVol(Prices!B2:B257)             annualised volatility from prices
'     =CAGR(Prices!B2:B257)               compound annual growth rate
'   Macros (Alt+F8):
'     RunFullReport     loops every ticker through the Dashboard, writes a
'                       "Report" sheet, highlights signals, exports a PDF
'     ImportPricesCSV   loads a Yahoo-style CSV (Date,...,Close) into Prices
'==============================================================================
Option Explicit

'---------------------------- UDFs -------------------------------------------
Public Function MaxDrawdown(prices As Range) As Double
    Dim peak As Double, dd As Double, worst As Double, c As Range
    peak = prices.Cells(1).Value
    For Each c In prices.Cells
        If IsNumeric(c.Value) And Not IsEmpty(c.Value) Then
            If c.Value > peak Then peak = c.Value
            dd = c.Value / peak - 1
            If dd < worst Then worst = dd
        End If
    Next c
    MaxDrawdown = worst
End Function

Public Function AnnVol(prices As Range, Optional periodsPerYear As Long = 252) As Double
    Dim n As Long, i As Long, rets() As Double
    n = prices.Cells.Count
    If n < 3 Then AnnVol = 0: Exit Function
    ReDim rets(1 To n - 1)
    For i = 2 To n
        rets(i - 1) = prices.Cells(i).Value / prices.Cells(i - 1).Value - 1
    Next i
    AnnVol = Application.WorksheetFunction.StDev(rets) * Sqr(periodsPerYear)
End Function

Public Function CAGR(prices As Range, Optional periodsPerYear As Long = 252) As Double
    Dim n As Long
    n = prices.Cells.Count
    CAGR = (prices.Cells(n).Value / prices.Cells(1).Value) ^ (periodsPerYear / (n - 1)) - 1
End Function

'---------------------------- Report macro ------------------------------------
Public Sub RunFullReport()
    Dim wsD As Worksheet, wsR As Worksheet, tickers As Variant
    Dim original As String, i As Long, r As Long

    Set wsD = ThisWorkbook.Worksheets("Dashboard")
    original = wsD.Range("C4").Value
    tickers = Array("0700.HK", "9988.HK", "0005.HK", "1299.HK", "0388.HK", "^HSI")

    Application.ScreenUpdating = False
    On Error GoTo CleanUp

    ' fresh Report sheet
    Application.DisplayAlerts = False
    On Error Resume Next: ThisWorkbook.Worksheets("Report").Delete: On Error GoTo CleanUp
    Application.DisplayAlerts = True
    Set wsR = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsR.Name = "Report"

    wsR.Range("A1").Value = "HK Equity Signal Report - generated " & Format(Now, "yyyy-mm-dd hh:nn")
    wsR.Range("A1").Font.Bold = True: wsR.Range("A1").Font.Size = 14
    wsR.Range("A3:I3").Value = Array("Ticker", "Company", "Last Close", "Period Return", _
        "Sharpe", "Max DD", "Signal", "MA Strategy", "Buy & Hold")
    With wsR.Range("A3:I3")
        .Font.Bold = True: .Font.Color = vbWhite: .Interior.Color = RGB(31, 56, 100)
    End With

    r = 4
    For i = LBound(tickers) To UBound(tickers)
        wsD.Range("C4").Value = tickers(i)          ' drive the dashboard
        Application.Calculate
        wsR.Cells(r, 1).Value = tickers(i)
        wsR.Cells(r, 2).Value = wsD.Range("F5").Value
        wsR.Cells(r, 3).Value = wsD.Range("F6").Value
        wsR.Cells(r, 4).Value = wsD.Range("F7").Value
        wsR.Cells(r, 5).Value = wsD.Range("F9").Value
        wsR.Cells(r, 6).Value = wsD.Range("F11").Value
        wsR.Cells(r, 7).Value = wsD.Range("I5").Value
        wsR.Cells(r, 8).Value = wsD.Range("I7").Value
        wsR.Cells(r, 9).Value = wsD.Range("I6").Value
        ' highlight signal
        If wsR.Cells(r, 7).Value = "BULLISH" Then
            wsR.Cells(r, 7).Interior.Color = RGB(198, 239, 206)
        Else
            wsR.Cells(r, 7).Interior.Color = RGB(255, 199, 206)
        End If
        r = r + 1
    Next i

    ' formatting
    wsR.Range("C4:C" & r - 1).NumberFormat = "#,##0.00"
    wsR.Range("D4:D" & r - 1 & ",F4:F" & r - 1 & ",H4:I" & r - 1).NumberFormat = "0.0%"
    wsR.Range("E4:E" & r - 1).NumberFormat = "0.00"
    wsR.Columns("A:I").AutoFit

    ' sort by Sharpe, best first
    wsR.Range("A3:I" & r - 1).Sort Key1:=wsR.Range("E3"), Order1:=xlDescending, Header:=xlYes

    ' export PDF next to the workbook
    wsR.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=ThisWorkbook.Path & Application.PathSeparator & "HK_Signal_Report_" & Format(Date, "yyyymmdd") & ".pdf"

CleanUp:
    wsD.Range("C4").Value = original                ' restore user's selection
    Application.ScreenUpdating = True
    Application.Calculate
    If Err.Number <> 0 Then
        MsgBox "Report failed: " & Err.Description, vbExclamation
    Else
        MsgBox "Report done: " & (r - 4) & " tickers. PDF saved next to the workbook.", vbInformation
    End If
End Sub

'---------------------------- CSV import --------------------------------------
' Pick a Yahoo Finance CSV (Date,Open,High,Low,Close,Adj Close,Volume) and
' write its Close column into the Prices sheet under the matching ticker header.
Public Sub ImportPricesCSV()
    Dim f As Variant, ticker As String, col As Variant, wb As Workbook
    Dim src As Worksheet, dst As Worksheet, lastR As Long, closeCol As Variant, i As Long, m As Variant

    f = Application.GetOpenFilename("CSV Files (*.csv), *.csv")
    If f = False Then Exit Sub
    ticker = InputBox("Ticker header in Prices (e.g. 0700.HK):", "Import", "0700.HK")
    If ticker = "" Then Exit Sub

    Set dst = ThisWorkbook.Worksheets("Prices")
    col = Application.Match(ticker, dst.Rows(1), 0)
    If IsError(col) Then MsgBox "Ticker not found in Prices header.": Exit Sub

    Set wb = Workbooks.Open(f)
    Set src = wb.Worksheets(1)
    closeCol = Application.Match("Close", src.Rows(1), 0)
    If IsError(closeCol) Then MsgBox "No 'Close' column in CSV.": wb.Close False: Exit Sub
    lastR = src.Cells(src.Rows.Count, 1).End(xlUp).Row

    ' match each date in Prices to the CSV and overwrite the close
    For i = 2 To dst.Cells(dst.Rows.Count, 1).End(xlUp).Row
        m = Application.Match(CLng(dst.Cells(i, 1).Value), src.Range("A2:A" & lastR), 0)
        If Not IsError(m) Then dst.Cells(i, col).Value = src.Cells(m + 1, closeCol).Value
    Next i
    wb.Close SaveChanges:=False
    Application.Calculate
    MsgBox "Imported " & ticker & ".", vbInformation
End Sub
