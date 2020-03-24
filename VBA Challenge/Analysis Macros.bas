Attribute VB_Name = "Module1"

Sub ClearOutputRangeMacro()

    Dim OutputRng As Range
    Set OutputRng = ActiveSheet.Columns("J:Z")
    OutputRng.Clear
    
End Sub

Sub GenericAnalysisMacro()

'DEFINE VARIABLES
    'Sheet and Functions
    Dim ws As Worksheet
    Dim wsf As Variant
    'Input/Output Containers - These are the ranges that contain Input/Output Columns
    Dim InputRng As Range, OutputRng As Range
    'Input Columns - I only need to use Ticker, Open, Close and Volume
    Dim InTickCol As Range, InOpenCol As Range, _
        InCloseCol As Range, InVolCol As Range
    'Output Columns - Ticker, Volume, Year Change, Percent Change
    Dim TickerCol As Range, VolumeCol As Range, _
        YChangeCol As Range, PChangeCol As Range
    'Sheet Constants - Define number of used rows and columns. Also count for all unique tickers
    Dim TotalRow As Long, TotalCol As Integer
    Dim UniqueTickerCount As Long
    'Loop Variables - I these are the temporary values as I iterate through the unique tickers
    Dim ticker As String, _
        firstrow As Long, lastrow As Long, _
        yearopen As Double, yearclose As Double 'First/Last Row a ticker appears in, its opening price and closing price
    
    
'CREATE WORKSHEET OBJECT AND FUNCTION NAMESPACE
    Set ws = Worksheets("A")
    Set wsf = ws.Application.WorksheetFunction
    'Size Active Cells
    TotalRow = ActiveSheet.UsedRange.Rows.Count
    TotalCol = 7
    
'CREATE RANGE OBJECTS'
    'Input/Output Containers
    Set InputRng = ws.Range(Rows(TotalRow), Columns(TotalCol))
    Set OutputRng = ws.Columns("J:Q")
    OutputRng.Clear
    'Input Ranges
    Set InTickCol = InputRng.Columns(1) 'Ticker
    Set InOpenCol = InputRng.Columns(3) 'Open
    Set InCloseCol = InputRng.Columns(6) 'Close
    Set InVolCol = InputRng.Columns(7) 'Volume
    'Output Ranges
    Set TickerCol = OutputRng.Columns(1) 'Ticker
    Set YChangeCol = OutputRng.Columns(2) 'Year Change (open - close)
    Set PChangeCol = OutputRng.Columns(3) 'Percentage Change (open-close)/open
    Set VolumeCol = OutputRng.Columns(4) 'Total Volume (Sum of daily volume)
    
'EXTRACT UNIQUE TICKERS
    'Copy Unique Tickers to New Range
    InputRng.Columns(1).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=TickerCol, Unique:=True
    'Find Number of unique stocks, (the subsequent loop is a simple counter of nonempty cells)
    UniqueTickerCount = 0
    For i = 1 To TotalRow
        If TickerCol.Cells(i, 1).Value <> "" Then
            UniqueTickerCount = UniqueTickerCount + 1
        End If
    Next i

'LABEL COLUMNS AND ANALYZE
    For i = 1 To UniqueTickerCount
        If i = 1 Then 'Label Row with Headers
            TickerCol.Cells(i, 1).Value = "Tickers"
            VolumeCol.Cells(i, 1).Value = "Total Volume"
            YChangeCol.Cells(i, 1).Value = "Year Change"
            PChangeCol.Cells(i, 1).Value = "Percent Change"
        Else 'Use formulae to extract information
            ticker = TickerCol.Cells(i, 1).Value 'Ticker String for this loop
            firstrow = wsf.Match(ticker, InTickCol, 0) 'First row "ticker" appears in
            lastrow = wsf.Match(ticker, InTickCol, 1)  'Last row  "ticker" appears in
            yearopen = wsf.Index(InOpenCol, firstrow, 1) 'Open price from first row
            yearclose = wsf.Index(InCloseCol, lastrow, 1) 'Close price from last row
            YChangeCol.Cells(i, 1).Value = yearclose - yearopen
            PChangeCol.Cells(i, 1).Value = (yearclose - yearopen) / yearopen
            VolumeCol.Cells(i, 1).Value = wsf.SumIf(InTickCol, ticker, InVolCol)
        End If
    Next i
        
'FORMAT COLUMNS
    YChangeCol.NumberFormat = "$#,##.00"
    PChangeCol.NumberFormat = "0.00%"
    
End Sub

