Attribute VB_Name = "Module1"
Sub GraphTwoStocks()

    'system setup
    Call SystemSetup
    
    'defining variables
    Dim cell As Range
    Dim i As Long
    Dim isFound1 As Boolean
    Dim isFound2 As Boolean
    Dim firstTime As Boolean
    Dim inp1 As String
    Dim inp2 As String
    Dim openCloseAdj As String
    Dim sym1 As String
    Dim sym1Open As Range
    Dim sym1Close As Range
    Dim sym1AdjClose As Range
    Dim sym1Date As Range
    Dim sym1Start As Long
    Dim sym1End As Long
    Dim sym1Range As Range
    Dim sym2 As String
    Dim sym2Open As Range
    Dim sym2Close As Range
    Dim sym2AdjClose As Range
    Dim sym2Date As Range
    Dim sym2Start As Long
    Dim sym2End As Long
    Dim sym2Range As Range
    
    'setting variables
    isFound1 = False
    isFound2 = False
    i = 0
    firstTime = True
    
    'user inputs symbol, matching symbols
    inp1 = InputBox("Enter company 1 symbol:", "Graphing Two Stocks")
    inp2 = InputBox("Enter company 2 symbol:", "Graphing Two Stocks")
    For Each cell In Worksheets("StockData").Range(Worksheets("StockData").Range("A2"), Worksheets("StockData").Range("A2").End(xlDown))
        If UCase(cell.Value) = UCase(inp1) Then
            isFound1 = True
            sym1 = cell.Value
        End If
        If UCase(cell.Value) = UCase(inp2) Then
            isFound2 = True
            sym2 = cell.Value
        End If
    Next cell

    'input not validated
    If isFound1 = False Or isFound2 = False Then
        If isFound1 = False Then
            Call MsgBox("Company symbol '" & inp1 & "' cannot be found.", vbOKOnly, "Error!")
        End If
        If isFound2 = False Then
            Call MsgBox("Company symbol '" & inp2 & "' cannot be found.", vbOKOnly, "Error!")
        End If
        Exit Sub
    End If
    
    'ask what they want to graph
    openCloseAdj = MsgBox("Would you like to graph the opening costs (Yes), the closing costs (No) or the adjusted closing costs (Cancel)?", vbYesNoCancel, "Opening, Closing, or Adjusted Closing Costs")

    'find symbol 1 range and associated ranges
    For Each cell In Range("Symbol")
        i = i + 1
        If cell.Value = sym1 And firstTime = True Then
            sym1Start = i
            firstTime = False
        End If
        If cell.Value <> sym1 And firstTime = False Then
            sym1End = i - 1
            Exit For
        End If
    Next cell
    Set sym1Range = Range(Worksheets("HistoryData").Range("A1").Offset(sym1Start, 0), Worksheets("HistoryData").Range("A1").Offset(sym1End, 0))
    Set sym1Date = sym1Range.Offset(0, 4)
    Set sym1Open = sym1Range.Offset(0, 5)
    Set sym1Close = sym1Range.Offset(0, 8)
    Set sym1AdjClose = sym1Range.Offset(0, 9)
    
    'find symbol 2 range and associated ranges
    i = 0
    firstTime = True
    For Each cell In Range("Symbol")
        i = i + 1
        If cell.Value = sym2 And firstTime = True Then
            sym2Start = i
            firstTime = False
        End If
        If cell.Value <> sym2 And firstTime = False Then
            sym2End = i - 1
            Exit For
        End If
    Next cell
    Set sym2Range = Range(Worksheets("HistoryData").Range("A1").Offset(sym2Start, 0), Worksheets("HistoryData").Range("A1").Offset(sym2End, 0))
    Set sym2Date = sym2Range.Offset(0, 4)
    Set sym2Open = sym2Range.Offset(0, 5)
    Set sym2Close = sym2Range.Offset(0, 8)
    Set sym2AdjClose = sym2Range.Offset(0, 9)
    
    'user wants opening costs
    If openCloseAdj = vbYes Then
        
        'set up chart data, title, axes, type
        Worksheets("HistoryData").Shapes.AddChart.Select
        With ActiveChart
            .SetSourceData Source:=Union(sym1Open, sym2Open, sym1Date, sym2Date)
            .ChartType = xlXYScatterSmooth
            .SetElement (msoElementChartTitleAboveChart)
            .ChartTitle.Text = sym1 & " vs " & sym2 & " Opening"
            .SetElement (msoElementPrimaryCategoryAxisTitleHorizontal)
            .Axes(xlCategory).AxisTitle.Text = "Date"
            .SetElement (msoElementPrimaryValueAxisTitleRotated)
            .Axes(xlValue).AxisTitle.Text = "Opening Costs"
            .SetElement (msoElementLegendRight)

            'organize symbol data
            .SeriesCollection("Series1").Delete
            With .SeriesCollection.NewSeries
                .XValues = sym1Date
                .Values = sym1Open
                .Name = sym1
            End With
            With .SeriesCollection.NewSeries
                .XValues = sym2Date
                .Values = sym2Open
                .Name = sym2
            End With
        
        'move and name sheet
            .Location (xlLocationAsNewSheet)
        End With
        ActiveSheet.Name = sym1 & "vs" & sym2 & "Opening"
    End If
    
    'user wants closing costs
    If openCloseAdj = vbNo Then
        
        'set up chart data, title, axes, type
        Worksheets("HistoryData").Shapes.AddChart.Select
        With ActiveChart
            .SetSourceData Source:=Union(sym1Close, sym2Close, sym1Date, sym2Date)
            .ChartType = xlXYScatterSmooth
            .SetElement (msoElementChartTitleAboveChart)
            .ChartTitle.Text = sym1 & " vs " & sym2 & " Closing"
            .SetElement (msoElementPrimaryCategoryAxisTitleHorizontal)
            .Axes(xlCategory).AxisTitle.Text = "Date"
            .SetElement (msoElementPrimaryValueAxisTitleRotated)
            .Axes(xlValue).AxisTitle.Text = "Closing Costs"
            .SetElement (msoElementLegendRight)

            'organize symbol data
            .SeriesCollection("Series1").Delete
            With .SeriesCollection.NewSeries
                .XValues = sym1Date
                .Values = sym1Close
                .Name = sym1
            End With
            With .SeriesCollection.NewSeries
                .XValues = sym2Date
                .Values = sym2Close
                .Name = sym2
            End With
        
        'move and name sheet
            .Location (xlLocationAsNewSheet)
        End With
        ActiveSheet.Name = sym1 & "vs" & sym2 & "Closing"
    End If
    
    'user wants adjusted closing costs
    If openCloseAdj = vbCancel Then
        
        'set up chart data, title, axes, type
        Worksheets("HistoryData").Shapes.AddChart.Select
        With ActiveChart
            .SetSourceData Source:=Union(sym1AdjClose, sym2AdjClose, sym1Date, sym2Date)
            .ChartType = xlXYScatterSmooth
            .SetElement (msoElementChartTitleAboveChart)
            .ChartTitle.Text = sym1 & " vs " & sym2 & " Adjusted Closing"
            .SetElement (msoElementPrimaryCategoryAxisTitleHorizontal)
            .Axes(xlCategory).AxisTitle.Text = "Date"
            .SetElement (msoElementPrimaryValueAxisTitleRotated)
            .Axes(xlValue).AxisTitle.Text = "Adjusted Closing Costs"
            .SetElement (msoElementLegendRight)
            
            'organize symbol data
            .SeriesCollection("Series1").Delete
            With .SeriesCollection.NewSeries
                .XValues = sym1Date
                .Values = sym1AdjClose
                .Name = sym1
            End With
            With .SeriesCollection.NewSeries
                .XValues = sym2Date
                .Values = sym2AdjClose
                .Name = sym2
            End With
            
        'move and name sheet
            .Location (xlLocationAsNewSheet)
        End With
        ActiveSheet.Name = sym1 & "vs" & sym2 & "AdjClosing"
    End If
End Sub

