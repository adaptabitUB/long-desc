from __future__ import annotations

from pathlib import Path


BASE_DIR = Path.cwd()
OUTPUT_BAS = BASE_DIR / "output" / "StatisticsSummaryMacro.bas"


def build_vba_module() -> str:
    return """Attribute VB_Name = \"MacroStatisticSummary\"
Option Explicit

Private Const OUTPUT_MARKER As String = "__VBA_STATISTICAL_SUMMARY__"

Public Sub StatisticalSummary_ActiveSheet()
    ProcessWorksheet ActiveSheet
End Sub

Public Sub StatisticalSummary_AllSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ProcessWorksheet ws
    Next ws
End Sub

Private Sub ProcessWorksheet(ByVal ws As Worksheet)
    On Error GoTo ErrHandler

    Dim chObj As ChartObject
    If ws.ChartObjects.Count = 0 Then Exit Sub

    Set chObj = ws.ChartObjects(1)

    Dim categories As Variant
    On Error Resume Next
    categories = GetCategories(chObj.Chart)
    On Error GoTo 0

    Dim globalVals As Collection
    Set globalVals = New Collection

    Dim perSeries As Object
    Set perSeries = CreateObject("Scripting.Dictionary")

    Dim perCategory As Object
    Set perCategory = CreateObject("Scripting.Dictionary")

    Dim xNums As Collection
    Dim yNums As Collection
    Set xNums = New Collection
    Set yNums = New Collection

    Dim s As Series
    Dim vals As Variant
    Dim i As Long
    Dim valsLBound As Long
    Dim numVal As Double
    Dim ok As Boolean
    Dim serieName As String
    Dim catKey As String
    Dim xVal As Double
    Dim xOk As Boolean

    On Error Resume Next
    For Each s In chObj.Chart.SeriesCollection
        vals = SafeArrayFromVariant(s.Values)
        serieName = CleanSeriesName(s.Name)

        If Not perSeries.Exists(serieName) Then
            perSeries.Add serieName, New Collection
        End If

        valsLBound = LBound(vals)
        For i = valsLBound To UBound(vals)
            ok = TryGetDouble(vals(i), numVal)
            If ok Then
                globalVals.Add numVal
                perSeries(serieName).Add numVal

                If IsArray(categories) Then
                    catKey = CategoryAt(categories, i, valsLBound)
                    If catKey <> "" Then
                        If Not perCategory.Exists(catKey) Then
                            perCategory.Add catKey, New Collection
                        End If
                        perCategory(catKey).Add numVal
                    End If

                    xOk = TryCategoryNumeric(categories, i, valsLBound, xVal)
                    If xOk Then
                        xNums.Add xVal
                        yNums.Add numVal
                    End If
                End If
            End If
        Next i
    Next s
    On Error GoTo 0

    Dim startRow As Long
    ClearOldOutput ws
    startRow = FindOutputStartRow(ws)
    WriteSummary ws, startRow, chObj.Chart, categories, perCategory, perSeries, globalVals, xNums, yNums

    Exit Sub

ErrHandler:
    On Error Resume Next
    MsgBox "Error processing sheet '" & ws.Name & "': " & Err.Description, vbExclamation
    On Error GoTo 0
End Sub

Private Sub WriteSummary(ByVal ws As Worksheet, ByVal startRow As Long, ByVal ch As Chart, _
                         ByVal categories As Variant, ByVal perCategory As Object, ByVal perSeries As Object, _
                         ByVal globalVals As Collection, ByVal xNums As Collection, ByVal yNums As Collection)

    Dim r As Long
    r = startRow

    ws.Cells(r, 1).Value = OUTPUT_MARKER
    ws.Cells(r, 2).Value = "Statistical summary of the chart"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 2).Font.Bold = True
    r = r + 1

    ws.Cells(r, 1).Value = "Type"
    ws.Cells(r, 2).Value = ChartTypeName(ch.ChartType)
    r = r + 1

    Dim xMin As Variant, xMax As Variant, xInterval As Variant, xUnitat As Variant
    Dim yMin As Variant, yMax As Variant, yInterval As Variant, yUnitat As Variant
    Dim y2Min As Variant, y2Max As Variant, y2Interval As Variant, y2Unitat As Variant

    ReadAxisInfo ch, xlCategory, xMin, xMax, xInterval, xUnitat
    ReadAxisInfo ch, xlValue, yMin, yMax, yInterval, yUnitat
    ReadSecondaryAxisInfo ch, y2Min, y2Max, y2Interval, y2Unitat

    ws.Cells(r, 1).Value = "x_min": ws.Cells(r, 2).Value = xMin: r = r + 1
    ws.Cells(r, 1).Value = "x_max": ws.Cells(r, 2).Value = xMax: r = r + 1
    ws.Cells(r, 1).Value = "x_interval": ws.Cells(r, 2).Value = xInterval: r = r + 1
    ws.Cells(r, 1).Value = "x_unit": ws.Cells(r, 2).Value = xUnitat: r = r + 1
    ws.Cells(r, 1).Value = "y_min": ws.Cells(r, 2).Value = yMin: r = r + 1
    ws.Cells(r, 1).Value = "y_max": ws.Cells(r, 2).Value = yMax: r = r + 1
    ws.Cells(r, 1).Value = "y_interval": ws.Cells(r, 2).Value = yInterval: r = r + 1
    ws.Cells(r, 1).Value = "y_unit": ws.Cells(r, 2).Value = yUnitat: r = r + 1
    ws.Cells(r, 1).Value = "y2_min": ws.Cells(r, 2).Value = y2Min: r = r + 1
    ws.Cells(r, 1).Value = "y2_max": ws.Cells(r, 2).Value = y2Max: r = r + 1
    ws.Cells(r, 1).Value = "y2_interval": ws.Cells(r, 2).Value = y2Interval: r = r + 1
    ws.Cells(r, 1).Value = "y2_unit": ws.Cells(r, 2).Value = y2Unitat: r = r + 1

    r = r + 1
    ws.Cells(r, 1).Value = "category_quantity"
    ws.Cells(r, 2).Value = perCategory.Count
    r = r + 1
    ws.Cells(r, 1).Value = "category_name_list"
    ws.Cells(r, 2).Value = CategoriesAsText(categories)
    r = r + 1
    ws.Cells(r, 1).Value = "series_quantity"
    ws.Cells(r, 2).Value = perSeries.Count
    r = r + 1
    ws.Cells(r, 1).Value = "series_name_list"
    ws.Cells(r, 2).Value = DictionaryKeysAsText(perSeries)
    r = r + 2

    ws.Cells(r, 1).Value = "Global_statistics"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1
    r = WriteStatsBlock(ws, r, globalVals)

    ws.Cells(r, 1).Value = "minimum_value_chart"
    ws.Cells(r, 2).Value = GlobalMin(globalVals)
    r = r + 1
    ws.Cells(r, 1).Value = "maximum_value_chart"
    ws.Cells(r, 2).Value = GlobalMax(globalVals)
    r = r + 1

    Dim corr As Variant
    corr = PearsonFromCollections(xNums, yNums)
    ws.Cells(r, 1).Value = "correlation_xy"
    ws.Cells(r, 2).Value = corr
    r = r + 2

    ws.Cells(r, 1).Value = "Statistics_by_series"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1

    Dim key As Variant
    For Each key In perSeries.Keys
        ws.Cells(r, 1).Value = "serie"
        ws.Cells(r, 2).Value = CStr(key)
        r = r + 1
        r = WriteStatsBlock(ws, r, perSeries(key))
        r = r + 1
    Next key

    ws.Cells(r, 1).Value = "Statistics_by_category"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1

    For Each key In perCategory.Keys
        ws.Cells(r, 1).Value = "category"
        ws.Cells(r, 2).Value = CStr(key)
        r = r + 1
        r = WriteStatsBlock(ws, r, perCategory(key))
        r = r + 1
    Next key

    ws.Columns("A:B").AutoFit
End Sub

Private Function WriteStatsBlock(ByVal ws As Worksheet, ByVal rowStart As Long, ByVal vals As Collection) As Long
    On Error Resume Next
    
    Dim r As Long
    r = rowStart

    ws.Cells(r, 1).Value = "n": ws.Cells(r, 2).Value = vals.Count: r = r + 1

    If vals.Count = 0 Then
        ws.Cells(r, 1).Value = "sum": ws.Cells(r, 2).Value = "": r = r + 1
        ws.Cells(r, 1).Value = "min": ws.Cells(r, 2).Value = "": r = r + 1
        ws.Cells(r, 1).Value = "max": ws.Cells(r, 2).Value = "": r = r + 1
        ws.Cells(r, 1).Value = "range": ws.Cells(r, 2).Value = "": r = r + 1
        ws.Cells(r, 1).Value = "average": ws.Cells(r, 2).Value = "": r = r + 1
        ws.Cells(r, 1).Value = "median": ws.Cells(r, 2).Value = "": r = r + 1
        ws.Cells(r, 1).Value = "mode": ws.Cells(r, 2).Value = "[]": r = r + 1
        ws.Cells(r, 1).Value = "standard_deviation": ws.Cells(r, 2).Value = "": r = r + 1
        ws.Cells(r, 1).Value = "coefficient_of_variation": ws.Cells(r, 2).Value = "": r = r + 1
        ws.Cells(r, 1).Value = "q1": ws.Cells(r, 2).Value = "": r = r + 1
        ws.Cells(r, 1).Value = "q2": ws.Cells(r, 2).Value = "": r = r + 1
        ws.Cells(r, 1).Value = "q3": ws.Cells(r, 2).Value = "": r = r + 1
        ws.Cells(r, 1).Value = "iqr": ws.Cells(r, 2).Value = "": r = r + 1
        ws.Cells(r, 1).Value = "p10": ws.Cells(r, 2).Value = "": r = r + 1
        ws.Cells(r, 1).Value = "p90": ws.Cells(r, 2).Value = "": r = r + 1
        ws.Cells(r, 1).Value = "outliers_count": ws.Cells(r, 2).Value = 0: r = r + 1
        ws.Cells(r, 1).Value = "outliers_values": ws.Cells(r, 2).Value = "[]": r = r + 1
        ws.Cells(r, 1).Value = "linear_trend": ws.Cells(r, 2).Value = "": r = r + 1
        WriteStatsBlock = r
        On Error GoTo 0
        Exit Function
    End If

    Dim arr() As Double
    arr = CollectionToArray(vals)

    Dim s As Double, mn As Double, mx As Double, m As Double, med As Double
    Dim q1 As Double, q3 As Double, iqr As Double, p10 As Double, p90 As Double
    Dim std As Double, cv As Variant, slope As Variant

    s = SumArray(arr)
    mn = MinArray(arr)
    mx = MaxArray(arr)
    m = s / (UBound(arr) - LBound(arr) + 1)
    med = Percentile(arr, 0.5)
    q1 = Percentile(arr, 0.25)
    q3 = Percentile(arr, 0.75)
    iqr = q3 - q1
    p10 = Percentile(arr, 0.1)
    p90 = Percentile(arr, 0.9)
    std = StdSample(arr)

    If m = 0 Then
        cv = Empty
    Else
        cv = std / m
    End If

    slope = LinearSlope(arr)

    Dim outVals As Collection
    Set outVals = IqrOutliers(arr, q1, q3)

    ws.Cells(r, 1).Value = "sum": ws.Cells(r, 2).Value = Round6(s): r = r + 1
    ws.Cells(r, 1).Value = "min": ws.Cells(r, 2).Value = Round6(mn): r = r + 1
    ws.Cells(r, 1).Value = "max": ws.Cells(r, 2).Value = Round6(mx): r = r + 1
    ws.Cells(r, 1).Value = "range": ws.Cells(r, 2).Value = Round6(mx - mn): r = r + 1
    ws.Cells(r, 1).Value = "average": ws.Cells(r, 2).Value = Round6(m): r = r + 1
    ws.Cells(r, 1).Value = "median": ws.Cells(r, 2).Value = Round6(med): r = r + 1
    ws.Cells(r, 1).Value = "mode": ws.Cells(r, 2).Value = ModesAsText(arr): r = r + 1
    ws.Cells(r, 1).Value = "standard_deviation": ws.Cells(r, 2).Value = Round6(std): r = r + 1
    ws.Cells(r, 1).Value = "coefficient_of_variation": ws.Cells(r, 2).Value = IIf(IsEmpty(cv), "", Round6(cv)): r = r + 1
    ws.Cells(r, 1).Value = "q1": ws.Cells(r, 2).Value = Round6(q1): r = r + 1
    ws.Cells(r, 1).Value = "q2": ws.Cells(r, 2).Value = Round6(med): r = r + 1
    ws.Cells(r, 1).Value = "q3": ws.Cells(r, 2).Value = Round6(q3): r = r + 1
    ws.Cells(r, 1).Value = "iqr": ws.Cells(r, 2).Value = Round6(iqr): r = r + 1
    ws.Cells(r, 1).Value = "p10": ws.Cells(r, 2).Value = Round6(p10): r = r + 1
    ws.Cells(r, 1).Value = "p90": ws.Cells(r, 2).Value = Round6(p90): r = r + 1
    ws.Cells(r, 1).Value = "outliers_count": ws.Cells(r, 2).Value = outVals.Count: r = r + 1
    ws.Cells(r, 1).Value = "outliers_values": ws.Cells(r, 2).Value = CollectionAsText(outVals): r = r + 1
    ws.Cells(r, 1).Value = "linear_trend": ws.Cells(r, 2).Value = IIf(IsEmpty(slope), "", Round6(slope)): r = r + 1

    WriteStatsBlock = r
    On Error GoTo 0
End Function

Private Function FindOutputStartRow(ByVal ws As Worksheet) As Long
    Dim lastDataRow As Long
    Dim lastChartRow As Long
    Dim ch As ChartObject

    lastDataRow = 1
    If WorksheetFunction.CountA(ws.Cells) > 0 Then
        lastDataRow = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    End If

    lastChartRow = 1
    For Each ch In ws.ChartObjects
        lastChartRow = WorksheetFunction.Max(lastChartRow, ch.BottomRightCell.Row)
    Next ch

    FindOutputStartRow = WorksheetFunction.Max(lastDataRow, lastChartRow) + 2
End Function

Private Sub ClearOldOutput(ByVal ws As Worksheet)
    Dim found As Range
    Set found = ws.Cells.Find(What:=OUTPUT_MARKER, LookIn:=xlValues, LookAt:=xlWhole)
    If found Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Rows.Count
    ws.Range(ws.Cells(found.Row, 1), ws.Cells(lastRow, 2)).ClearContents
End Sub

Private Function SafeArrayFromVariant(ByVal valuesIn As Variant) As Variant
    Dim result() As Variant
    Dim i As Long

    If IsArray(valuesIn) Then
        SafeArrayFromVariant = valuesIn
    Else
        ReDim result(1 To 1)
        result(1) = valuesIn
        SafeArrayFromVariant = result
    End If
End Function

Private Function GetCategories(ByVal ch As Chart) As Variant
    Dim cats As Variant
    Dim result() As Variant
    Dim i As Long

    On Error Resume Next
    If ch.SeriesCollection.Count > 0 Then
        cats = ch.SeriesCollection(1).XValues
    End If
    On Error GoTo 0

    If IsEmpty(cats) Then
        ReDim result(1 To 0)
        GetCategories = result
        Exit Function
    End If

    If IsArray(cats) Then
        GetCategories = cats
    ElseIf IsObject(cats) Then
        If TypeName(cats) = "Range" Then
            If cats.Cells.Count = 0 Then
                ReDim result(1 To 0)
            Else
                ReDim result(1 To cats.Cells.Count)
                For i = 1 To cats.Cells.Count
                    result(i) = cats.Cells(i).Value
                Next i
            End If
            GetCategories = result
        Else
            ReDim result(1 To 1)
            result(1) = CStr(cats)
            GetCategories = result
        End If
    Else
        ReDim result(1 To 1)
        result(1) = cats
        GetCategories = result
    End If
End Function

Private Function CategoryAt(ByVal categories As Variant, ByVal idx As Long, ByVal refLBound As Long) As String
    Dim catIdx As Long
    On Error GoTo Fail
    
    If Not IsArray(categories) Then
        CategoryAt = ""
        Exit Function
    End If
    
    If UBound(categories) < LBound(categories) Then
        CategoryAt = ""
        Exit Function
    End If
    
    catIdx = LBound(categories) + (idx - refLBound)
    
    If catIdx < LBound(categories) Or catIdx > UBound(categories) Then
        CategoryAt = ""
        Exit Function
    End If
    
    CategoryAt = CStr(categories(catIdx))
    Exit Function
Fail:
    CategoryAt = ""
End Function

Private Function TryCategoryNumeric(ByVal categories As Variant, ByVal idx As Long, ByVal refLBound As Long, ByRef outVal As Double) As Boolean
    Dim catIdx As Long
    On Error GoTo Fail
    
    If Not IsArray(categories) Then
        TryCategoryNumeric = False
        Exit Function
    End If
    
    If UBound(categories) < LBound(categories) Then
        TryCategoryNumeric = False
        Exit Function
    End If
    
    catIdx = LBound(categories) + (idx - refLBound)
    
    If catIdx < LBound(categories) Or catIdx > UBound(categories) Then
        TryCategoryNumeric = False
        Exit Function
    End If
    
    TryCategoryNumeric = TryGetDouble(categories(catIdx), outVal)
    Exit Function
Fail:
    TryCategoryNumeric = False
End Function

Private Sub ReadAxisInfo(ByVal ch As Chart, ByVal axisType As XlAxisType, _
                         ByRef outMin As Variant, ByRef outMax As Variant, _
                         ByRef outInterval As Variant, ByRef outUnit As Variant)
    On Error GoTo NoAxis
    Dim ax As Axis
    Set ax = ch.Axes(axisType)

    If axisType = xlCategory Then
        outMin = AxisCategoryMin(ch)
        outMax = AxisCategoryMax(ch)
        outInterval = ax.TickLabelSpacing
    Else
        outMin = ax.MinimumScale
        outMax = ax.MaximumScale
        outInterval = ax.MajorUnit
    End If

    If ax.HasTitle Then
        outUnit = ax.AxisTitle.Text
    Else
        outUnit = ""
    End If
    Exit Sub

NoAxis:
    outMin = Empty
    outMax = Empty
    outInterval = Empty
    outUnit = ""
End Sub

Private Sub ReadSecondaryAxisInfo(ByVal ch As Chart, ByRef outMin As Variant, ByRef outMax As Variant, _
                                  ByRef outInterval As Variant, ByRef outUnit As Variant)
    On Error GoTo NoAxis
    Dim ax As Axis
    Set ax = ch.Axes(xlValue, xlSecondary)
    outMin = ax.MinimumScale
    outMax = ax.MaximumScale
    outInterval = ax.MajorUnit
    If ax.HasTitle Then
        outUnit = ax.AxisTitle.Text
    Else
        outUnit = ""
    End If
    Exit Sub
NoAxis:
    outMin = Empty
    outMax = Empty
    outInterval = Empty
    outUnit = ""
End Sub

Private Function AxisCategoryMin(ByVal ch As Chart) As Variant
    Dim c As Variant
    c = GetCategories(ch)
    On Error GoTo Fail
    AxisCategoryMin = c(LBound(c))
    Exit Function
Fail:
    AxisCategoryMin = Empty
End Function

Private Function AxisCategoryMax(ByVal ch As Chart) As Variant
    Dim c As Variant
    c = GetCategories(ch)
    On Error GoTo Fail
    AxisCategoryMax = c(UBound(c))
    Exit Function
Fail:
    AxisCategoryMax = Empty
End Function

Private Function CleanSeriesName(ByVal s As String) As String
    If Len(Trim$(s)) = 0 Then
        CleanSeriesName = "Value"
    ElseIf Left$(s, 1) = "=" Then
        CleanSeriesName = Replace$(s, "=", "")
    Else
        CleanSeriesName = s
    End If
End Function

Private Function TryGetDouble(ByVal v As Variant, ByRef outVal As Double) As Boolean
    On Error GoTo Fail
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        TryGetDouble = False
        Exit Function
    End If
    If IsNumeric(v) Then
        outVal = CDbl(v)
        TryGetDouble = True
        Exit Function
    End If
    TryGetDouble = False
    Exit Function
Fail:
    TryGetDouble = False
End Function

Private Function CollectionToArray(ByVal c As Collection) As Double()
    Dim arr() As Double
    Dim i As Long
    If c.Count = 0 Then
        ReDim arr(1 To 1)
        arr(1) = 0
        CollectionToArray = arr
        Exit Function
    End If
    ReDim arr(1 To c.Count)
    For i = 1 To c.Count
        arr(i) = CDbl(c(i))
    Next i
    CollectionToArray = arr
End Function

Private Function SumArray(ByRef arr() As Double) As Double
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        SumArray = SumArray + arr(i)
    Next i
End Function

Private Function MinArray(ByRef arr() As Double) As Double
    Dim i As Long
    If UBound(arr) < LBound(arr) Then
        MinArray = 0
        Exit Function
    End If
    MinArray = arr(LBound(arr))
    For i = LBound(arr) + 1 To UBound(arr)
        If arr(i) < MinArray Then MinArray = arr(i)
    Next i
End Function

Private Function MaxArray(ByRef arr() As Double) As Double
    Dim i As Long
    If UBound(arr) < LBound(arr) Then
        MaxArray = 0
        Exit Function
    End If
    MaxArray = arr(LBound(arr))
    For i = LBound(arr) + 1 To UBound(arr)
        If arr(i) > MaxArray Then MaxArray = arr(i)
    Next i
End Function

Private Function Percentile(ByRef arr() As Double, ByVal p As Double) As Double
    Dim n As Long
    n = UBound(arr) - LBound(arr) + 1
    
    If n <= 0 Then
        Percentile = 0
        Exit Function
    End If
    
    Dim sorted() As Double
    sorted = SortedCopy(arr)

    If n = 1 Then
        Percentile = sorted(LBound(sorted))
        Exit Function
    End If

    Dim pos As Double
    pos = (n - 1) * p

    Dim lowIdx As Long, highIdx As Long
    lowIdx = Int(pos)
    highIdx = lowIdx
    If pos > lowIdx Then highIdx = lowIdx + 1
    
    lowIdx = WorksheetFunction.Max(0, WorksheetFunction.Min(lowIdx, n - 1))
    highIdx = WorksheetFunction.Max(0, WorksheetFunction.Min(highIdx, n - 1))

    Dim baseIdx As Long
    baseIdx = LBound(sorted)

    If lowIdx = highIdx Then
        Percentile = sorted(baseIdx + lowIdx)
    Else
        Dim w As Double
        w = pos - lowIdx
        Percentile = sorted(baseIdx + lowIdx) * (1# - w) + sorted(baseIdx + highIdx) * w
    End If
End Function

Private Function SortedCopy(ByRef arr() As Double) As Double()
    Dim result() As Double
    result = arr
    
    Dim n As Long
    n = UBound(result) - LBound(result) + 1

    If n <= 1 Then
        SortedCopy = result
        Exit Function
    End If

    Dim i As Long, j As Long, tmp As Double
    For i = LBound(result) To UBound(result) - 1
        For j = i + 1 To UBound(result)
            If result(j) < result(i) Then
                tmp = result(i)
                result(i) = result(j)
                result(j) = tmp
            End If
        Next j
    Next i
    SortedCopy = result
End Function

Private Function StdSample(ByRef arr() As Double) As Double
    Dim n As Long
    n = UBound(arr) - LBound(arr) + 1

    If n <= 1 Then
        StdSample = 0
        Exit Function
    End If

    Dim m As Double
    If n > 0 Then
        m = SumArray(arr) / n
    End If

    Dim i As Long
    Dim acc As Double
    For i = LBound(arr) To UBound(arr)
        acc = acc + (arr(i) - m) * (arr(i) - m)
    Next i

    On Error Resume Next
    StdSample = Sqr(acc / (n - 1))
    On Error GoTo 0
End Function

Private Function ModesAsText(ByRef arr() As Double) As String
    Dim n As Long
    n = UBound(arr) - LBound(arr) + 1
    
    If n <= 0 Then
        ModesAsText = "[]"
        Exit Function
    End If
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    Dim key As String
    On Error Resume Next
    For i = LBound(arr) To UBound(arr)
        key = CStr(arr(i))
        If dict.Exists(key) Then
            dict(key) = dict(key) + 1
        Else
            dict.Add key, 1
        End If
    Next i
    On Error GoTo 0

    Dim maxCount As Long
    maxCount = 1
    Dim k As Variant
    For Each k In dict.Keys
        If dict(k) > maxCount Then maxCount = dict(k)
    Next k

    If maxCount <= 1 Then
        ModesAsText = "[]"
        Exit Function
    End If

    Dim out As String
    out = "["
    For Each k In dict.Keys
        If dict(k) = maxCount Then
            If out <> "[" Then out = out & ", "
            out = out & CStr(Round6(CDbl(k)))
        End If
    Next k
    out = out & "]"

    ModesAsText = out
End Function

Private Function LinearSlope(ByRef arr() As Double) As Variant
    Dim n As Long
    n = UBound(arr) - LBound(arr) + 1
    If n < 2 Then
        LinearSlope = Empty
        Exit Function
    End If

    Dim xMean As Double
    xMean = (n - 1) / 2#

    On Error Resume Next
    Dim yMean As Double
    yMean = SumArray(arr) / n
    
    Dim i As Long
    Dim dx As Double
    Dim num As Double
    Dim den As Double

    For i = 0 To n - 1
        dx = i - xMean
        num = num + dx * (arr(LBound(arr) + i) - yMean)
        den = den + dx * dx
    Next i

    If den = 0 Then
        LinearSlope = Empty
    Else
        LinearSlope = num / den
    End If
    
    On Error GoTo 0
End Function

Private Function IqrOutliers(ByRef arr() As Double, ByVal q1 As Double, ByVal q3 As Double) As Collection
    Dim result As New Collection
    Dim lowFence As Double
    Dim highFence As Double
    lowFence = q1 - 1.5 * (q3 - q1)
    highFence = q3 + 1.5 * (q3 - q1)

    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) < lowFence Or arr(i) > highFence Then
            result.Add Round6(arr(i))
        End If
    Next i

    Set IqrOutliers = result
End Function

Private Function CollectionAsText(ByVal c As Collection) As String
    Dim i As Long
    Dim t As String
    t = "["
    For i = 1 To c.Count
        If i > 1 Then t = t & ", "
        t = t & CStr(c(i))
    Next i
    t = t & "]"
    CollectionAsText = t
End Function

Private Function CategoriesAsText(ByVal categories As Variant) As String
    On Error GoTo Fail
    If Not IsArray(categories) Then
        CategoriesAsText = ""
        Exit Function
    End If
    If UBound(categories) < LBound(categories) Then
        CategoriesAsText = ""
        Exit Function
    End If
    Dim i As Long
    Dim t As String
    t = ""
    For i = LBound(categories) To UBound(categories)
        If i > LBound(categories) Then t = t & "; "
        t = t & CStr(categories(i))
    Next i
    CategoriesAsText = t
    Exit Function
Fail:
    CategoriesAsText = ""
End Function

Private Function DictionaryKeysAsText(ByVal dict As Object) As String
    Dim k As Variant
    Dim t As String
    t = ""
    For Each k In dict.Keys
        If Len(t) > 0 Then t = t & "; "
        t = t & CStr(k)
    Next k
    DictionaryKeysAsText = t
End Function

Private Function GlobalMin(ByVal vals As Collection) As Variant
    If vals.Count = 0 Then
        GlobalMin = Empty
        Exit Function
    End If
    Dim arr() As Double
    arr = CollectionToArray(vals)
    GlobalMin = Round6(MinArray(arr))
End Function

Private Function GlobalMax(ByVal vals As Collection) As Variant
    If vals.Count = 0 Then
        GlobalMax = Empty
        Exit Function
    End If
    Dim arr() As Double
    arr = CollectionToArray(vals)
    GlobalMax = Round6(MaxArray(arr))
End Function

Private Function PearsonFromCollections(ByVal xs As Collection, ByVal ys As Collection) As Variant
    If xs.Count <> ys.Count Or xs.Count < 2 Then
        PearsonFromCollections = Empty
        Exit Function
    End If

    Dim i As Long
    Dim xm As Double, ym As Double
    For i = 1 To xs.Count
        xm = xm + CDbl(xs(i))
        ym = ym + CDbl(ys(i))
    Next i
    xm = xm / xs.Count
    ym = ym / ys.Count

    Dim num As Double
    Dim denx As Double
    Dim deny As Double
    Dim dx As Double
    Dim dy As Double

    For i = 1 To xs.Count
        dx = CDbl(xs(i)) - xm
        dy = CDbl(ys(i)) - ym
        num = num + dx * dy
        denx = denx + dx * dx
        deny = deny + dy * dy
    Next i

    If denx = 0 Or deny = 0 Then
        PearsonFromCollections = Empty
    Else
        PearsonFromCollections = Round6(num / (Sqr(denx) * Sqr(deny)))
    End If
End Function

Private Function Round6(ByVal v As Double) As Double
    Round6 = WorksheetFunction.Round(v, 6)
End Function

Private Function ChartTypeName(ByVal ct As XlChartType) As String
    Select Case ct
        Case xl3DArea : ChartTypeName = "3D Area"
        Case xl3DAreaStacked : ChartTypeName = "3D Stacked Area"
        Case xl3DAreaStacked100 : ChartTypeName = "3D 100% Stacked Area"
        Case xl3DBarClustered : ChartTypeName = "3D Clustered Bar"
        Case xl3DBarStacked : ChartTypeName = "3D Stacked Bar"
        Case xl3DBarStacked100 : ChartTypeName = "3D 100% Stacked Bar"
        Case xl3DColumn : ChartTypeName = "3D Column"
        Case xl3DColumnClustered : ChartTypeName = "3D Clustered Column"
        Case xl3DColumnStacked : ChartTypeName = "3D Stacked Column"
        Case xl3DColumnStacked100 : ChartTypeName = "3D 100% Stacked Column"
        Case xl3DLine : ChartTypeName = "3D Line"
        Case xl3DPie : ChartTypeName = "3D Pie"
        Case xl3DPieExploded : ChartTypeName = "Exploded 3D Pie"
        Case xlArea : ChartTypeName = "Area"
        Case xlAreaEx : ChartTypeName = "Area Ex"
        Case xlAreaStacked : ChartTypeName = "Stacked Area"
        Case xlAreaStacked100 : ChartTypeName = "100% Stacked Area"
        Case xlAreaStacked100Ex : ChartTypeName = "100% Stacked Area Ex"
        Case xlAreaStackedEx : ChartTypeName = "Stacked Area Ex"
        Case xlBarClustered : ChartTypeName = "Clustered Bar"
        Case xlBarClusteredEx : ChartTypeName = "Clustered Bar Ex"
        Case xlBarOfPie : ChartTypeName = "Bar of Pie"
        Case xlBarStacked : ChartTypeName = "Stacked Bar"
        Case xlBarStacked100 : ChartTypeName = "100% Stacked Bar"
        Case xlBarStacked100Ex : ChartTypeName = "100% Stacked Bar Ex"
        Case xlBarStackedEx : ChartTypeName = "Stacked Bar Ex"
        Case xlBoxwhisker : ChartTypeName = "Box & Whisker"
        Case xlBubble : ChartTypeName = "Bubble"
        Case xlBubble3DEffect : ChartTypeName = "Bubble with 3D effects"
        Case xlBubbleEx : ChartTypeName = "Bubble Ex"
        Case xlColumnClustered : ChartTypeName = "Clustered Column"
        Case xlColumnClusteredEx : ChartTypeName = "Clustered Column Ex"
        Case xlColumnStacked : ChartTypeName = "Stacked Column"
        Case xlColumnStacked100 : ChartTypeName = "100% Stacked Column"
        Case xlColumnStacked100Ex : ChartTypeName = "100% Stacked Column Ex"
        Case xlColumnStackedEx : ChartTypeName = "Stacked Column Ex"
        Case xlCombo : ChartTypeName = "Custom Combination"
        Case xlComboAreaStackedColumnClustered : ChartTypeName = "Stacked Area - Clustered Column"
        Case xlComboColumnClusteredLine : ChartTypeName = "Clustered Column - Line"
        Case xlComboColumnClusteredLineSecondaryAxis : ChartTypeName = "Clustered Column - Line on Secondary Axis"
        Case xlConeBarClustered : ChartTypeName = "Clustered Cone Bar"
        Case xlConeBarStacked : ChartTypeName = "Stacked Cone Bar"
        Case xlConeBarStacked100 : ChartTypeName = "100% Stacked Cone Bar"
        Case xlConeCol : ChartTypeName = "3D Cone Column"
        Case xlConeColClustered : ChartTypeName = "Clustered Cone Column"
        Case xlConeColStacked : ChartTypeName = "Stacked Cone Column"
        Case xlConeColStacked100 : ChartTypeName = "100% Stacked Cone Column"
        Case xlCylinderBarClustered : ChartTypeName = "Clustered Cylinder Bar"
        Case xlCylinderBarStacked : ChartTypeName = "Stacked Cylinder Bar"
        Case xlCylinderBarStacked100 : ChartTypeName = "100% Stacked Cylinder Bar"
        Case xlCylinderCol : ChartTypeName = "3D Cylinder Column"
        Case xlCylinderColClustered : ChartTypeName = "Clustered Cylinder Column"
        Case xlCylinderColStacked : ChartTypeName = "Stacked Cylinder Column"
        Case xlCylinderColStacked100 : ChartTypeName = "100% Stacked Cylinder Column"
        Case xlDoughnut : ChartTypeName = "Doughnut"
        Case xlDoughnutEx : ChartTypeName = "Doughnut Ex"
        Case xlDoughnutExploded : ChartTypeName = "Exploded Doughnut"
        Case xlFunnel : ChartTypeName = "Funnel"
        Case xlHistogram : ChartTypeName = "Histogram"
        Case xlLine : ChartTypeName = "Line"
        Case xlLineEx : ChartTypeName = "Line Ex"
        Case xlLineMarkers : ChartTypeName = "Line with Markers"
        Case xlLineMarkersStacked : ChartTypeName = "Stacked Line with Markers"
        Case xlLineMarkersStacked100 : ChartTypeName = "100% Stacked Line with Markers"
        Case xlLineStacked : ChartTypeName = "Stacked Line"
        Case xlLineStacked100 : ChartTypeName = "100% Stacked Line"
        Case xlLineStacked100Ex : ChartTypeName = "100% Stacked Line Ex"
        Case xlLineStackedEx : ChartTypeName = "Stacked Line Ex"
        Case xlOtherCombinations : ChartTypeName = "Other Combinations"
        Case xlPareto : ChartTypeName = "Pareto"
        Case xlPie : ChartTypeName = "Pie"
        Case xlPieEx : ChartTypeName = "Pie Ex"
        Case xlPieExploded : ChartTypeName = "Exploded Pie"
        Case xlPieOfPie : ChartTypeName = "Pie of Pie"
        Case xlPyramidBarClustered : ChartTypeName = "Clustered Pyramid Bar"
        Case xlPyramidBarStacked : ChartTypeName = "Stacked Pyramid Bar"
        Case xlPyramidBarStacked100 : ChartTypeName = "100% Stacked Pyramid Bar"
        Case xlPyramidCol : ChartTypeName = "3D Pyramid Column"
        Case xlPyramidColClustered : ChartTypeName = "Clustered Pyramid Column"
        Case xlPyramidColStacked : ChartTypeName = "Stacked Pyramid Column"
        Case xlPyramidColStacked100 : ChartTypeName = "100% Stacked Pyramid Column"
        Case xlRadar : ChartTypeName = "Radar"
        Case xlRadarFilled : ChartTypeName = "Filled Radar"
        Case xlRadarMarkers : ChartTypeName = "Radar with Data Markers"
        Case xlRegionMap : ChartTypeName = "Map chart"
        Case xlStockHLC : ChartTypeName = "High-Low-Close"
        Case xlStockOHLC : ChartTypeName = "Open-High-Low-Close"
        Case xlStockVHLC : ChartTypeName = "Volume-High-Low-Close"
        Case xlStockVOHLC : ChartTypeName = "Volume-Open-High-Low-Close"
        Case xlSuggestedChart : ChartTypeName = "Suggested"
        Case xlSunburst : ChartTypeName = "Sunburst"
        Case xlSurface : ChartTypeName = "3D Surface"
        Case xlSurfaceTopView : ChartTypeName = "Surface (Top View)"
        Case xlSurfaceTopViewWireframe : ChartTypeName = "Surface (Top View wireframe)"
        Case xlSurfaceWireframe : ChartTypeName = "3D Surface (wireframe)"
        Case xlTreemap : ChartTypeName = "Treemap"
        Case xlWaterfall : ChartTypeName = "Waterfall"
        Case xlXYScatter : ChartTypeName = "Scatter"
        Case xlXYScatterEx : ChartTypeName = "Scatter Ex"
        Case xlXYScatterLines : ChartTypeName = "Scatter with Lines"
        Case xlXYScatterLinesNoMarkers : ChartTypeName = "Scatter with Lines and No Data Markers"
        Case xlXYScatterSmooth : ChartTypeName = "Scatter with Smoothed Lines"
        Case xlXYScatterSmoothNoMarkers : ChartTypeName = "Scatter with Smoothed Lines and No Data Markers"
        
        ' Special cases
        Case -4111: ChartTypeName = "None"
        Case -4151: ChartTypeName = "Automatic"
        
        Case Else
            ChartTypeName = "ChartType_" & CStr(ct)
    End Select
End Function
"""


def main() -> None:
    OUTPUT_BAS.parent.mkdir(parents=True, exist_ok=True)
    vba_code = build_vba_module()
    OUTPUT_BAS.write_text(vba_code, encoding="utf-8")

    print(f"VBA file generated: {OUTPUT_BAS}")
    print("Import this .bas within the Excel's VBA editor (ALT+F11 > File > Import File).")
    print("Execute for the current sheet or for the whole sheets.")


if __name__ == "__main__":
    main()