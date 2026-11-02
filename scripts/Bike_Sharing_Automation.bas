Attribute VB_Name = "BikeSharingAutomation"
Option Explicit

'Import into a macro-enabled copy (.xlsm) of the report using Alt+F11 > File > Import File.
'Run RefreshBikeSharingReport after replacing or appending rows in the Clean Data sheet.
'Works on Windows desktop Excel. Does not import the three source files or train an AI model.
Public Sub RefreshBikeSharingReport()
    Dim ws As Worksheet, daily As Worksheet, seg As Worksheet, insight As Worksheet
    Dim lastRow As Long, r As Long, n As Long, dateKey As String, k As Variant
    Dim dates As Object, hours As Object, weather As Object, dayType As Object, seasons As Object, months As Object
    Dim counts As Object, rec As Variant, demand As Double, casual As Double, registered As Double
    Dim dt As Date, h As String, w As String, typ As String, season As String, m As String
    Dim sorted As Variant, i As Long, j As Long, temp As Variant, avg As Double, sd As Double, z As Double
    Dim alpha As Double, level As Double, total As Double, sq As Double, days As Long, nextDate As Date
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Set ws = ThisWorkbook.Worksheets("Clean Data")
    Set daily = ThisWorkbook.Worksheets("Daily Trends")
    Set seg = ThisWorkbook.Worksheets("Demand Segments")
    Set insight = ThisWorkbook.Worksheets("Insights & Forecast")
    Set dates = CreateObject("Scripting.Dictionary")
    Set hours = CreateObject("Scripting.Dictionary")
    Set weather = CreateObject("Scripting.Dictionary")
    Set dayType = CreateObject("Scripting.Dictionary")
    Set seasons = CreateObject("Scripting.Dictionary")
    Set months = CreateObject("Scripting.Dictionary")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Err.Raise vbObjectError + 10, , "No data in Clean Data"
    For r = 2 To lastRow
        If IsDate(ws.Cells(r, 2).Value) And IsNumeric(ws.Cells(r, 16).Value) And IsNumeric(ws.Cells(r, 14).Value) And IsNumeric(ws.Cells(r, 15).Value) Then
            demand = CDbl(ws.Cells(r, 16).Value)
            casual = CDbl(ws.Cells(r, 14).Value)
            registered = CDbl(ws.Cells(r, 15).Value)
            If Abs(demand - casual - registered) < 0.001 And demand >= 0 Then
                dt = CDate(ws.Cells(r, 2).Value)
                dateKey = Format$(dt, "yyyy-mm-dd")
                h = CStr(ws.Cells(r, 6).Value)
                w = CStr(ws.Cells(r, 19).Value)
                typ = CStr(ws.Cells(r, 17).Value)
                season = CStr(ws.Cells(r, 18).Value)
                m = CStr(ws.Cells(r, 5).Value)
                AddAgg dates, dateKey, demand
                AddAgg hours, h, demand
                AddAgg weather, w, demand
                AddAgg dayType, typ, demand
                AddAgg seasons, season, demand
                AddAgg months, m, demand
            End If
        End If
    Next r
    If dates.Count = 0 Then Err.Raise vbObjectError + 11, , "No valid rows"
    daily.Range("A2:E100000").ClearContents
    sorted = dates.Keys
    For i = LBound(sorted) To UBound(sorted) - 1
        For j = i + 1 To UBound(sorted)
            If CStr(sorted(i)) > CStr(sorted(j)) Then temp = sorted(i): sorted(i) = sorted(j): sorted(j) = temp
        Next j
    Next i
    total = 0: sq = 0: days = dates.Count
    For i = 0 To days - 1
        rec = dates(sorted(i))
        total = total + rec(0)
        sq = sq + rec(0) * rec(0)
    Next i
    avg = total / days
    If days > 1 Then sd = Sqr(Application.Max(0, (sq - days * avg * avg) / (days - 1))) Else sd = 0
    alpha = 0.35
    For i = 0 To days - 1
        rec = dates(sorted(i))
        daily.Cells(i + 2, 1).Value = sorted(i)
        daily.Cells(i + 2, 2).Value = rec(0)
        daily.Cells(i + 2, 3).Value = rec(1)
        If sd > 0 Then z = (rec(0) - avg) / sd Else z = 0
        daily.Cells(i + 2, 4).Value = z
        daily.Cells(i + 2, 5).Value = IIf(Abs(z) >= 2, "REVIEW", "Normal")
        If i = 0 Then level = rec(0) Else level = alpha * rec(0) + (1 - alpha) * level
    Next i
    seg.Range("A2:AC100000").ClearContents
    WriteAgg seg, hours, 1
    WriteAgg seg, weather, 6
    WriteAgg seg, dayType, 11
    WriteAgg seg, seasons, 21
    WriteAgg seg, months, 26
    'Registered/casual segmentation is calculated from valid source rows.
    Dim sumCasual As Double, sumRegistered As Double, validN As Long
    For r = 2 To lastRow
        If IsNumeric(ws.Cells(r, 16).Value) And IsNumeric(ws.Cells(r, 14).Value) And IsNumeric(ws.Cells(r, 15).Value) Then
            If Abs(CDbl(ws.Cells(r, 16).Value) - CDbl(ws.Cells(r, 14).Value) - CDbl(ws.Cells(r, 15).Value)) < 0.001 Then
                sumCasual = sumCasual + CDbl(ws.Cells(r, 14).Value)
                sumRegistered = sumRegistered + CDbl(ws.Cells(r, 15).Value)
                validN = validN + 1
            End If
        End If
    Next r
    seg.Cells(2, 16).Value = "Casual": seg.Cells(2, 17).Value = sumCasual: seg.Cells(2, 18).Value = sumCasual / Application.Max(1, validN): seg.Cells(2, 19).Value = validN
    seg.Cells(3, 16).Value = "Registered": seg.Cells(3, 17).Value = sumRegistered: seg.Cells(3, 18).Value = sumRegistered / Application.Max(1, validN): seg.Cells(3, 19).Value = validN
    nextDate = DateAdd("d", 1, CDate(sorted(days - 1)))
    For i = 0 To 6
        insight.Cells(9 + i, 1).Value = Format$(DateAdd("d", i, nextDate), "yyyy-mm-dd")
        insight.Cells(9 + i, 2).Value = Round(level, 0)
        insight.Cells(9 + i, 3).Value = Application.Max(0, Round(level - 1.96 * sd, 0))
        insight.Cells(9 + i, 4).Value = Round(level + 1.96 * sd, 0)
    Next i
    ThisWorkbook.RefreshAll
    Application.CalculateFull
    MsgBox "Daily trends, segment summaries, anomalies and baseline forecast refreshed.", vbInformation
CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    MsgBox "Refresh failed: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Private Sub AddAgg(ByRef dict As Object, ByVal key As String, ByVal value As Double)
    Dim a As Variant
    If dict.Exists(key) Then
        a = dict(key): a(0) = a(0) + value: a(1) = a(1) + 1
    Else
        a = Array(0#, 0&): a(0) = value: a(1) = 1
    End If
    dict(key) = a
End Sub

Private Sub WriteAgg(ByRef sh As Worksheet, ByRef dict As Object, ByVal col As Long)
    Dim k As Variant, a As Variant, r As Long
    r = 2
    For Each k In dict.Keys
        a = dict(k)
        sh.Cells(r, col).Value = k
        sh.Cells(r, col + 1).Value = a(0)
        sh.Cells(r, col + 2).Value = a(0) / a(1)
        sh.Cells(r, col + 3).Value = a(1)
        r = r + 1
    Next k
End Sub
