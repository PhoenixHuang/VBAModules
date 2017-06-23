Option Explicit
Sub Main()
Call Step1_GetRawdata
Call Step2_GetEntityAndEntitled
Call Step3_GetSolidRawData
Call Step4_GetRidOfDuplicate
Call Step5_GetReport
Application.DisplayAlerts = False
Application.ActiveWorkbook.Worksheets("Rawdata").Delete
Application.ActiveWorkbook.Worksheets("NewRawdata").Delete
Application.DisplayAlerts = True
End Sub

Private Sub Step1_GetRawdata()
Application.ScreenUpdating = False
Dim wb As Workbook
Set wb = Application.ActiveWorkbook

Dim AttWs As Worksheet
Set AttWs = wb.Worksheets("May - time & attendence report")
Dim RawDataWs As Worksheet
Set RawDataWs = wb.Worksheets.Add
RawDataWs.Name = "RawData"

RawDataWs.Cells(1, 1) = "Name"
RawDataWs.Cells(1, 2) = "BadgeID"
RawDataWs.Cells(1, 3) = "OutDate"
RawDataWs.Cells(1, 4) = "OutTime"
RawDataWs.Cells(1, 5) = "Entity"
RawDataWs.Cells(1, 6) = "CostCenter"
RawDataWs.Cells(1, 7) = "Entitled"
RawDataWs.Cells(1, 8) = "WorkingDay"

Dim iLastRow As Long
iLastRow = AttWs.Cells.SpecialCells(xlCellTypeLastCell).Row

Dim i As Long
Dim iRaw As Long
iRaw = 2
For i = 2 To iLastRow
    If AttWs.Cells(i, 1).MergeArea.Cells(1, 1) <> "" Then
        Dim j As Integer
        For j = 1 To 200
            If AttWs.Cells(i + j, 1).MergeArea.Cells(1, 1) <> "" Then
                Exit For
            End If
            If AttWs.Cells(i + j, 1).MergeArea.Cells(1, 1) = "" Then
                If AttWs.Cells(i + j, 10).MergeArea.Cells(1, 1) <> "" And Len(AttWs.Cells(i + j, 10).MergeArea.Cells(1, 1)) < 25 Then
                    RawDataWs.Cells(iRaw, 1) = AttWs.Cells(i, 1)
                    RawDataWs.Cells(iRaw, 2) = AttWs.Cells(i + j, 10).MergeArea.Cells(1, 1)
                    RawDataWs.Cells(iRaw, 3) = AttWs.Cells(i + j, 13).MergeArea.Cells(1, 1)
                    RawDataWs.Cells(iRaw, 4) = AttWs.Cells(i + j, 16).MergeArea.Cells(1, 1)
                    iRaw = iRaw + 1
                End If
            End If
        Next
    End If
Next
RawDataWs.Columns(3).NumberFormatLocal = "yyyy/m/d"
RawDataWs.Columns(4).NumberFormatLocal = "h:mm:ss;@"
Application.ScreenUpdating = True
End Sub

Private Sub Step2_GetEntityAndEntitled()
Application.ScreenUpdating = False
Dim wb As Workbook
Set wb = Application.ActiveWorkbook

Dim AttWs As Worksheet
Set AttWs = wb.Worksheets("Access card register")

Dim RawDataWs As Worksheet
Set RawDataWs = wb.Worksheets("RawData")

Dim i As Long
Dim iLastRow As Long
iLastRow = RawDataWs.Cells.SpecialCells(xlCellTypeLastCell).Row

For i = 2 To iLastRow
    RawDataWs.Cells(i, 5).FormulaR1C1 = "=VLOOKUP(RC[-3],'Access card register'!C[-4]:C,3,FALSE)"
    RawDataWs.Cells(i, 6).FormulaR1C1 = "=VLOOKUP(RC[-4],'Access card register'!C[-5]:C,4,FALSE)"
    RawDataWs.Cells(i, 7).FormulaR1C1 = "=VLOOKUP(RC[-5],'Access card register'!C[-6]:C,5,FALSE)"
    RawDataWs.Cells(i, 8).FormulaR1C1 = "=VLOOKUP(RC[-5],'Working day'!C[-6]:C[-5],2,FALSE)"
Next
Application.ScreenUpdating = True
End Sub

Private Sub Step3_GetSolidRawData()
Application.ScreenUpdating = False
Dim wb As Workbook
Set wb = Application.ActiveWorkbook

Dim RawDataWs As Worksheet
Set RawDataWs = wb.Worksheets("RawData")

Dim NewRawDataWs As Worksheet
Set NewRawDataWs = wb.Worksheets.Add
NewRawDataWs.Name = "NewRawData"

Dim i As Long
Dim iLastRow As Long
iLastRow = RawDataWs.Cells.SpecialCells(xlCellTypeLastCell).Row
Dim iLastCol As Long
iLastCol = RawDataWs.Cells.SpecialCells(xlCellTypeLastCell).Column

Dim j As Long
For i = 1 To iLastRow
        For j = 1 To iLastCol
            NewRawDataWs.Cells(i, j) = RawDataWs.Cells(i, j).Text
        Next
Next
NewRawDataWs.Cells.Replace What:="0", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
NewRawDataWs.Cells.Replace What:="#N/A", Replacement:="", LookAt:=xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

Dim iRow As Long
For iRow = iLastRow To 2 Step -1
    If NewRawDataWs.Cells(iRow, 7) <> "Y" Or NewRawDataWs.Cells(iRow, 8) <> "working day" Or IsNotOT(CDbl(RawDataWs.Cells(iRow, 4).Value)) Then
       NewRawDataWs.Rows(iRow).Delete
    End If
Next

Application.ScreenUpdating = True
End Sub

Private Sub Step4_GetRidOfDuplicate()
Application.ScreenUpdating = False
Dim wb As Workbook
Set wb = Application.ActiveWorkbook

Dim RawDataWs As Worksheet
Set RawDataWs = wb.Worksheets("NewRawData")

'sort
RawDataWs.Sort.SortFields.Clear
RawDataWs.Sort.SortFields.Add Key:=RawDataWs.Columns(2), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
RawDataWs.Sort.SortFields.Add Key:=RawDataWs.Columns(3), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
RawDataWs.Sort.SortFields.Add Key:=RawDataWs.Columns(4), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
RawDataWs.Sort.SetRange RawDataWs.Cells
RawDataWs.Sort.Header = xlYes
RawDataWs.Sort.MatchCase = False
RawDataWs.Sort.Orientation = xlTopToBottom
RawDataWs.Sort.SortMethod = xlPinYin
RawDataWs.Sort.Apply
    
Dim i As Long
Dim iLastRow As Long
iLastRow = RawDataWs.Cells.SpecialCells(xlCellTypeLastCell).Row
Dim iLastCol As Long
iLastCol = RawDataWs.Cells.SpecialCells(xlCellTypeLastCell).Column

'remove same id duplicate day records,let one day have only one record
For i = iLastRow To 2 Step -1
    If RawDataWs.Cells(i, 2) = RawDataWs.Cells(i - 1, 2) And RawDataWs.Cells(i, 3) = RawDataWs.Cells(i - 1, 3) Then
        RawDataWs.Rows(i).Delete
    End If
Next

'if id is the same, date diff is one day, previous day is after 10PM, next day is before 6AM, then remove one
For i = iLastRow To 2 Step -1
    If RawDataWs.Cells(i, 2) <> "" And RawDataWs.Cells(i, 2) = RawDataWs.Cells(i - 1, 2) Then
        Dim df As Integer
        df = DateDiff("d", RawDataWs.Cells(i, 3).Value, RawDataWs.Cells(i - 1, 3))
        If df = -1 And DateDiff("s", RawDataWs.Cells(i - 1, 4), 0.916666667) < 0 And DateDiff("s", RawDataWs.Cells(i, 4), 0.25) > 0 Then
            RawDataWs.Rows(i).Delete
        End If
    End If
Next
Application.ScreenUpdating = True
End Sub
Private Sub Step5_GetReport()
Application.ScreenUpdating = False
Dim wb As Workbook
Set wb = Application.ActiveWorkbook

Dim RawDataWs As Worksheet
Set RawDataWs = wb.Worksheets("NewRawData")

Dim reportws As Worksheet
Set reportws = wb.Worksheets.Add
reportws.Name = "Report"

reportws.Cells(1, 1) = "Entity"
reportws.Cells(1, 2) = "Badge ID"
reportws.Cells(1, 3) = "Staff name"
reportws.Cells(1, 4) = "Cost center"
reportws.Cells(1, 5) = "No. of days entitled to meal allowance"
reportws.Cells(1, 6) = "Total Amount (HK$)"

Dim imonth As Integer
Dim iyear As Integer
iyear = Year(RawDataWs.Cells(2, 3))
imonth = Month(RawDataWs.Cells(2, 3))
Dim imonthstart, imonthend
imonthstart = DateSerial(iyear, imonth, 1)
imonthend = DateSerial(iyear, imonth + 1, 1) - 1

Dim j As Integer
For j = 7 To imonthend - imonthstart + 7
    reportws.Cells(1, j) = DateSerial(iyear, imonth, 1 + j - 7)
    reportws.Cells(1, j).NumberFormatLocal = "m/d"
    reportws.Cells(2, j) = WeekdayName(Weekday(reportws.Cells(1, j), vbSunday))
Next
        
Dim i As Long
Dim iLastRow As Long
iLastRow = RawDataWs.Cells.SpecialCells(xlCellTypeLastCell).Row
Dim iLastCol As Long
iLastCol = RawDataWs.Cells.SpecialCells(xlCellTypeLastCell).Column

Dim iRow As Long
iRow = 3
For i = 2 To iLastRow
    If RawDataWs.Cells(i, 1) <> "" Then
        If RawDataWs.Cells(i, 1) <> reportws.Cells(iRow, 3) Then
            iRow = iRow + 1
            reportws.Cells(iRow, 1) = RawDataWs.Cells(i, 5)
            reportws.Cells(iRow, 2) = RawDataWs.Cells(i, 2)
            reportws.Cells(iRow, 3) = RawDataWs.Cells(i, 1)
            reportws.Cells(iRow, 4) = RawDataWs.Cells(i, 6)
            reportws.Cells(iRow, 5) = 1
            reportws.Cells(iRow, 6) = 50
            reportws.Cells(iRow, 6 + Day(RawDataWs.Cells(i, 3))) = 1
        Else
            reportws.Cells(iRow, 5) = reportws.Cells(iRow, 5) + 1
            reportws.Cells(iRow, 6) = reportws.Cells(iRow, 6) + 50
            reportws.Cells(iRow, 6 + Day(RawDataWs.Cells(i, 3))) = 1
        End If
    End If
Next
reportws.Rows(1).Font.Bold = True
reportws.Rows(2).Font.Bold = True
reportws.Range("G:AK").HorizontalAlignment = xlCenter
reportws.Range("A:F").EntireColumn.AutoFit
Application.ScreenUpdating = True
End Sub

Public Function IsNotOT(s As Double) As Boolean
Dim b As Long
b = DateDiff("s", s, 0.916666667)
Dim c As Long
c = DateDiff("s", s, 0.25)
If b < 0 Or c > 0 Then
    IsNotOT = False
Else
    IsNotOT = True
End If
End Function
