Attribute VB_Name = "obpi"

Function goalset(s, k, r, sigma, t, z)
'
' Inputs are S = initial stock price
'            K = strike price
'            r = risk-free rate
'            sigma = volatility
'            q = dividend yield
'            T = time to maturity
'
Dim d1, d2, N1, N2

    d1 = (Log(s / k) + (r + 0.5 * sigma * sigma) * t) / (sigma * Sqr(t))
    d2 = d1 - sigma * Sqr(t)
    N1 = Application.NormSDist(d1)
    N2 = Application.NormSDist(d2)
    BSC = s * N1 - Exp(-r * t) * k * N2
    goalset = BSC / k + Exp(-r * t) - 1 / z
    
End Function

Sub obpi_backtest()

'Dim s, r, k, sigma, t, d1, d2, N1, N2, w1, w2
Dim wb1 As Workbook, wb2 As Workbook, wb3 As Workbook, wb4 As Workbook, wb5 As Workbook
Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet, ws4 As Worksheet, ws5 As Worksheet, ws6 As Worksheet
Dim ws7 As Worksheet, ws8 As Worksheet, ws9 As Worksheet
Dim rng As Range, rng1 As Range, y As Integer
Dim kd As Double, s As Double, rf As Double, sigma As Double, z As Double, t As Double

Set ws1 = ThisWorkbook.Worksheets("control")
Set ws2 = ThisWorkbook.Worksheets("OBPI")
Set ws3 = ThisWorkbook.Worksheets("OBPI_strategy")

ws2.Range("a3:f65536").ClearContents
'ws2.Columns("L:V").Delete Shift:=xlToLeft
ws2.Columns("l:v").ClearContents
ws3.Range("a3:e65536").ClearContents

'yr = ws1.Range("c10")        ''''year
yr = 2015                    '''the backtest from 2015 backwark 5 years
cd = ws1.Range("c7")         ''''code
sea = 4                      '''season

ws2.Activate

rw1 = ws2.Range("l65536").End(xlUp).Row
ws2.Range("l" & rw1).Select

Do
    rw1 = ws2.Range("l65536").End(xlUp).Row    'this sentence is for the beginning of the loop

    ws2.Range("l" & rw1).Select
    
     Set ie = CreateObject("InternetExplorer.Application")
        ie.Navigate "http://quotes.money.163.com/trade/lsjysj_" & ws1.Range("c7") & ".html?year=" & yr & "&season=" & sea & ""
        ie.Visible = False
        Do While ie.Busy Or ie.ReadyState <> 4: DoEvents: Loop
            Set AA = ie.Document.getelementsbytagname("table")
            With ws2
                k = ws2.Range("l65536").End(xlUp).Row
            
                For i = 0 To AA(3).Rows.Length - 1
                    For jj = 0 To AA(3).Rows(i).Cells.Length - 1
                        .Cells(k, jj + 12) = AA(3).Rows(i).Cells(jj).INNERTEXT
                    Next
                    k = k + 1
                 Next
            End With
        ie.Quit
     
     
    ws2.Range("l" & rw1 & ":v" & rw1).Delete Shift:=xlUp
    
    If sea = 1 Then
        yr = yr - 1
        sea = sea + 4
    End If
    
    sea = sea - 1

Loop Until Application.WorksheetFunction.CountA(ws2.Range("l1" & ":l" & rw1)) >= 252


'''''''''''''''''below is processing data

r1 = ws2.Range("l65536").End(xlUp).Row

    Range("l1:v1").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("OBPI").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("OBPI").Sort.SortFields.Add Key:=Range("l1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("OBPI").Sort
        .SetRange Range("l1" & ":v" & r1)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
rw1 = ws2.Range("a65536").End(xlUp).Row
rw2 = ws2.Range("l65536").End(xlUp).Row

ws2.Range("a3" & ":a" & rw2 + 1).Value = ws2.Range("l2" & ":l" & rw2 + 1).Value
ws2.Range("b3" & ":b" & rw2 + 1).Value = ws2.Range("p2" & ":p" & rw2 + 1).Value

For i = 4 To rw2 + 1
    ws2.Cells(i, 3).Value = Application.WorksheetFunction.Ln(ws2.Cells(i, 2).Value / ws2.Cells(i - 1, 2).Value)
Next

For j = 33 To rw2 + 1
    ws2.Cells(j, 5).Value = Sqr(Application.WorksheetFunction.Var(ws2.Range("c" & j - 29 & ":c" & j)))
Next

    b = Application.WorksheetFunction.Count(ws2.Range("e33" & ":e" & rw2 + 1))
For k = 33 To rw2 + 1
    a = Application.WorksheetFunction.Count(ws2.Range("e33" & ":e" & rw2 + 1))
    ws2.Cells(k, 6).Value = b / a
    b = b - 1
Next

'''fill the data of rf to the sheet of OBPI

r2 = ws2.Range("a65536").End(xlUp).Row

For i = 3 To r2
    If ws2.Cells(i, 1).Value < CDate("2007/7/23") Then
        ws2.Cells(i, 4).Value = 0.0072
    ElseIf CDate("2007/12/21") > ws2.Cells(i, 1).Value >= CDate("2007/7/23") Then
        ws2.Cells(i, 4).Value = 0.0081
    ElseIf CDate("2008/12/29") > ws2.Cells(i, 1).Value >= CDate("2007/12/21") Then
        ws2.Cells(i, 4).Value = 0.0072
    ElseIf CDate("2011/2/9") > ws2.Cells(i, 1).Value >= CDate("2008/12/29") Then
        ws2.Cells(i, 4).Value = 0.0036
    ElseIf CDate("2011/4/6") > ws2.Cells(i, 1).Value >= CDate("2011/2/9") Then
        ws2.Cells(i, 4).Value = 0.004
    ElseIf CDate("2012/6/7") > ws2.Cells(i, 1).Value >= CDate("2011/4/6") Then
        ws2.Cells(i, 4).Value = 0.005
    ElseIf CDate("2012/7/6") > ws2.Cells(i, 1).Value >= CDate("2012/6/7") Then
        ws2.Cells(i, 4).Value = 0.004
    ElseIf ws2.Cells(i, 1).Value >= CDate("2012/7/6") Then
        ws2.Cells(i, 4).Value = 0.0035
    End If
Next
'''calculate the delevery price
kd = ws2.Cells(33, 9).Value
s = ws2.Cells(33, 2).Value
rf = ws2.Cells(33, 3).Value
t = ws2.Cells(33, 5).Value
sigma = ws2.Cells(33, 4).Value * Sqr(252)   '''the sigma should be annualized
z = ws2.Cells(33, 6).Value

ws2.Cells(33, 8).GoalSeek Goal:=0, ChangingCell:=ws2.Cells(33, 9)

''''''''''''''''''''''''''''''''it is the presentation of the sheet of OBPI_strategy below

ws3.Range("a1:a65536").Value = ws2.Range("a1:a65536").Value   '''date
ws3.Range("b1:b65536").Value = ws2.Range("b1:b65536").Value   '''close_price


r1 = ws2.Range("a65536").End(xlUp).Row
For i = 33 To r1
    s = ws2.Cells(i, 2).Value
    rf = ws2.Cells(i, 4).Value
    sigma = ws2.Cells(i, 5).Value * Sqr(252)
    t = ws2.Cells(i, 6).Value
    k = ws2.Cells(33, 9).Value
    d1 = (Log(s / k) + (rf + 0.5 * sigma * sigma) * t) / (sigma * Sqr(t))
    d2 = d1 - sigma * Sqr(t)
    minus_d2 = sigma * Sqr(t) - d1
    
    N1 = Application.NormSDist(d1)
    N2 = Application.NormSDist(minus_d2)
    
    ws3.Cells(i, 3).Value = s * N1 / (s * N1 + k * Exp(-rf * t) * N2)                   '''the calculation of w1
    ws3.Cells(i, 4).Value = k * Exp(-rf * t) * N2 / (s * N1 + k * Exp(-rf * t) * N2)     '''the calculation of w2
Next
    
ws3.Cells(33, 5).Value = 1
For j = 34 To r1
    r2 = ws3.Range("a65536").End(xlUp).Row
    t = 1 / Application.WorksheetFunction.Count(ws3.Range("c33" & ":c" & r2))     '''this variable should be adjust(should be decreasing)
    rf = ws2.Cells(j - 1, 4).Value
'    t = ws2.Cells(j - 1, 6).Value
    v1 = ws3.Cells(j - 1, 5).Value * ws3.Cells(j - 1, 3).Value / ws3.Cells(j - 1, 2).Value * ws3.Cells(j, 2).Value
    v2 = ws3.Cells(j - 1, 5).Value * ws3.Cells(j - 1, 4).Value * Exp(rf * t)
    ws3.Cells(j, 5).Value = v1 + v2
Next


End Sub


Sub monte_carlo_simulation()

'Dim s, r, k, sigma, t, d1, d2, N1, N2, w1, w2
Dim wb1 As Workbook, wb2 As Workbook, wb3 As Workbook, wb4 As Workbook, wb5 As Workbook
Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet, ws4 As Worksheet, ws5 As Worksheet, ws6 As Worksheet
Dim ws7 As Worksheet, ws8 As Worksheet, ws9 As Worksheet
Dim rng As Range, rng1 As Range, y As Integer
Dim kd As Double, s As Double, rf As Double, sigma As Double, z As Double, t As Double

Set ws1 = ThisWorkbook.Worksheets("control")
Set ws2 = ThisWorkbook.Worksheets("OBPI")
Set ws3 = ThisWorkbook.Worksheets("OBPI_strategy")

yr = ws1.Range("c10")        ''''year
cd = ws1.Range("c7")         ''''code
sea = 4                      '''season

ws2.Range("x1:aj65536").ClearContents

ws2.Activate

rw1 = ActiveSheet.Range("x65536").End(xlUp).Row

Do
    Set ie = CreateObject("InternetExplorer.Application")
        ie.Navigate "http://quotes.money.163.com/trade/lsjysj_" & ws1.Range("c7") & ".html?year=" & yr & "&season=" & sea & ""
        ie.Visible = False
        Do While ie.Busy Or ie.ReadyState <> 4: DoEvents: Loop
            Set AA = ie.Document.getelementsbytagname("table")
            With ws2
                k = ws2.Range("x65536").End(xlUp).Row
            
                For i = 0 To AA(3).Rows.Length - 1
                    For jj = 0 To AA(3).Rows(i).Cells.Length - 1
                        .Cells(k, jj + 24) = AA(3).Rows(i).Cells(jj).INNERTEXT
                    Next
                    k = k + 1
                 Next
            End With
        ie.Quit
     
     
    ws2.Range("x" & rw1 & ":ah" & rw1).Delete Shift:=xlUp
    
    If sea = 1 Then
        yr = yr - 1
        sea = sea + 4
    End If
    
    sea = sea - 1
    
    rw1 = ws2.Range("x65536").End(xlUp).Row    'this sentence is for the beginning of the loop

    
Loop Until Application.WorksheetFunction.CountA(ws2.Range("x1" & ":x" & rw1)) >= 1000

    Range("X2:AH2").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("OBPI").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("OBPI").Sort.SortFields.Add Key:=Range("X2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("OBPI").Sort
        .SetRange Range("X2:AH1500")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

''''''''to get the current stock price

s = ws1.Cells(9, 5).Value

'''to get the average mu & sigma & dt

rw2 = ws2.Range("ab65536").End(xlUp).Row
For i = 1 To rw2 - 1
    If ws2.Cells(i, 28).Value <> "" Then
        rw3 = ws2.Cells(i, 28).Row
        ws2.Cells(i + 1, 35).Value = Log(ws2.Cells(i + 1, 28).Value / ws2.Cells(i, 28).Value)
    End If
Next
mu = Application.WorksheetFunction.Average(ws2.Range("ai" & rw3 & ":ai" & rw2))

'sigma = Sqr(Application.WorksheetFunction.Var(ws2.Range("ai" & rw3 & ":ai" & rw2)))
'sigma = 0.1

ws2.Activate
    SolverOk SetCell:=Range("am1002"), MaxMinVal:=1, ValueOf:=0, ByChange:=Range("ao3:ao5") _
        , Engine:=1, EngineDesc:="GRG Nonlinear"
    SolverSolve userfinish = False
Application.DisplayAlerts = False
sigma = ws2.Cells(6, 42).Value

dt = 1 / 252

ws3.Cells(3, 14).Value = s
For j = 4 To 255
    drift = Application.WorksheetFunction.NormInv(Rnd(), 0, 1)
    ws3.Cells(j, 14).Value = ws3.Cells(j - 1, 14).Value + (ws3.Cells(j - 1, 14).Value * mu * dt + ws3.Cells(j - 1, 14).Value * sigma * drift * Sqr(dt))
Next
    
'''''''''''run w1& w2 and the total value

ws2.Range("as1:at253").Value = ws3.Range("m3:n255").Value

For i = 2 To 253
    ws2.Cells(i, 47).Value = Log(ws2.Cells(i, 46).Value / ws2.Cells(i - 1, 46).Value)
Next

'For i = 31 To 253   '''the method of moving-average volatility
'    ws2.Cells(i, 48).Value = Sqr(Application.WorksheetFunction.Var(ws2.Range("au" & i - 30 & ":au" & i)))
'Next

ws2.Cells(31, 52).GoalSeek Goal:=0, ChangingCell:=ws2.Cells(31, 53)  '''get k

r1 = ws2.Range("as65536").End(xlUp).Row
For i = 1 To r1
    s = ws2.Cells(i, 46).Value
    rf = ws2.Cells(31, 49).Value                 ''''to implement the data from website
    k = ws2.Cells(31, 53).Value
'    sigma = ws2.Cells(i, 40).Value
'    sigma = 0.1
    sigma = ws2.Cells(6, 42).Value
    t = 1
    d1 = (Log(s / k) + (rf + 0.5 * sigma * sigma) * t) / (sigma * Sqr(t))
    d2 = d1 - sigma * Sqr(t)
    minus_d2 = sigma * Sqr(t) - d1
    
    N1 = Application.NormSDist(d1)
    N2 = Application.NormSDist(minus_d2)
    
    ws3.Cells(i + 2, 15).Value = s * N1 / (s * N1 + k * Exp(-rf * t) * N2)                 '''the calculation of w1
    ws3.Cells(i + 2, 16).Value = k * Exp(-rf * t) * N2 / (s * N1 + k * Exp(-rf * t) * N2)   '''the calculation of w2
    t = t - 1 / 253
Next
    
ws3.Cells(3, 17).Value = 1
For j = 4 To r1 + 2
    t = 1 / 223
    rf = ws2.Cells(31, 49).Value
    v1 = ws3.Cells(j - 1, 17).Value * ws3.Cells(j - 1, 15).Value / ws3.Cells(j - 1, 14).Value * ws3.Cells(j, 14).Value
    v2 = ws3.Cells(j - 1, 17).Value * ws3.Cells(j - 1, 16).Value * Exp(rf * t)
    ws3.Cells(j, 17).Value = v1 + v2
Next


End Sub

Sub refresh_trend_of_stock()

Dim wb1 As Workbook, wb2 As Workbook, wb3 As Workbook, wb4 As Workbook, wb5 As Workbook
Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet, ws4 As Worksheet, ws5 As Worksheet, ws6 As Worksheet
Dim ws7 As Worksheet, ws8 As Worksheet, ws9 As Worksheet
Dim rng As Range, rng1 As Range, y As Integer
Dim kd As Double, s As Double, rf As Double, sigma As Double, z As Double, t As Double

Set ws1 = ThisWorkbook.Worksheets("control")
Set ws2 = ThisWorkbook.Worksheets("OBPI")
Set ws3 = ThisWorkbook.Worksheets("OBPI_strategy")



'i = 0
'Do
'    i = i + 1
'    If ws2.Cells(i, 28).Value <> "" Then
'        s = ws2.Cells(i, 28).Value
'    End If
'Loop Until ws2.Cells(i, 28).Value <> ""
s = ws1.Cells(9, 5).Value

'''to get the average mu & sigma & dt

rw2 = ws2.Range("ab65536").End(xlUp).Row
For i = 1 To rw2 - 1
    If ws2.Cells(i, 28).Value <> "" Then
        rw3 = ws2.Cells(i, 28).Row
        ws2.Cells(i + 1, 35).Value = Log(ws2.Cells(i + 1, 28).Value / ws2.Cells(i, 28).Value)
    End If
Next
mu = Application.WorksheetFunction.Average(ws2.Range("ai" & rw3 & ":ai" & rw2))

'sigma = Sqr(Application.WorksheetFunction.Var(ws2.Range("ai" & rw3 & ":ai" & rw2)))
'sigma = 0.1

ws2.Activate
    SolverOk SetCell:=Range("am1002"), MaxMinVal:=1, ValueOf:=0, ByChange:=Range("ao3:ao5") _
        , Engine:=1, EngineDesc:="GRG Nonlinear"
    SolverSolve userfinish = False
Application.DisplayAlerts = False
sigma = ws2.Cells(6, 42).Value

dt = 1 / 252

ws3.Cells(3, 14).Value = s
For j = 4 To 255
    drift = Application.WorksheetFunction.NormInv(Rnd(), 0, 1)
    ws3.Cells(j, 14).Value = ws3.Cells(j - 1, 14).Value + (ws3.Cells(j - 1, 14).Value * mu * dt + ws3.Cells(j - 1, 14).Value * sigma * drift * Sqr(dt))
Next
    
'''''''''''run w1& w2 and the total value

ws2.Range("as1:at253").Value = ws3.Range("m3:n255").Value

For i = 2 To 253
    ws2.Cells(i, 47).Value = Log(ws2.Cells(i, 46).Value / ws2.Cells(i - 1, 46).Value)
Next

'For i = 31 To 253   '''the method of moving-average volatility
'    ws2.Cells(i, 48).Value = Sqr(Application.WorksheetFunction.Var(ws2.Range("au" & i - 30 & ":au" & i)))
'Next

ws2.Cells(31, 52).GoalSeek Goal:=0, ChangingCell:=ws2.Cells(31, 53)  '''get k

r1 = ws2.Range("as65536").End(xlUp).Row
For i = 1 To r1
    s = ws2.Cells(i, 46).Value
    rf = ws2.Cells(31, 49).Value                 ''''to implement the data from website
    k = ws2.Cells(31, 53).Value
'    sigma = ws2.Cells(i, 40).Value
'    sigma = 0.1
    sigma = ws2.Cells(6, 42).Value
    t = 1
    d1 = (Log(s / k) + (rf + 0.5 * sigma * sigma) * t) / (sigma * Sqr(t))
    d2 = d1 - sigma * Sqr(t)
    minus_d2 = sigma * Sqr(t) - d1
    
    N1 = Application.NormSDist(d1)
    N2 = Application.NormSDist(minus_d2)
    
    ws3.Cells(i + 2, 15).Value = s * N1 / (s * N1 + k * Exp(-rf * t) * N2)                 '''the calculation of w1
    ws3.Cells(i + 2, 16).Value = k * Exp(-rf * t) * N2 / (s * N1 + k * Exp(-rf * t) * N2)   '''the calculation of w2
    t = t - 1 / 253
Next
    
ws3.Cells(3, 17).Value = 1
For j = 4 To r1 + 2
    t = 1 / 223
    rf = ws2.Cells(31, 49).Value
    v1 = ws3.Cells(j - 1, 17).Value * ws3.Cells(j - 1, 15).Value / ws3.Cells(j - 1, 14).Value * ws3.Cells(j, 14).Value
    v2 = ws3.Cells(j - 1, 17).Value * ws3.Cells(j - 1, 16).Value * Exp(rf * t)
    ws3.Cells(j, 17).Value = v1 + v2
Next




End Sub

