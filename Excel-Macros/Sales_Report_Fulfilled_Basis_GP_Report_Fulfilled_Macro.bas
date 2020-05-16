Attribute VB_Name = "Module13"
Sub SalesReportFulfilledBasisGPReportFulfilled()
Attribute SalesReportFulfilledBasisGPReportFulfilled.VB_ProcData.VB_Invoke_Func = "F\n14"
'
' SalesReportFulfilledBasis Macro
'

'
'Validation Test
    'Test to see if the sheets have the correct name.
    result = MsgBox("Please ensure that you have Set up your Spread Sheet Correctly." _
    & vbNewLine & vbNewLine & _
    "HOW DO YOU SET IT UP?" & vbNewLine & _
    Chr(149) & " Download ""Gross Profit Report - Fulfilled"" from REX" & vbNewLine & _
    "    " & Chr(176) & " Filters for atleast 1 week to 1 month. Ideally, 1-4 weeks M-F" & vbNewLine & _
    "    " & Chr(176) & " Downlaod as CSV" & vbNewLine & _
    Chr(149) & " Download Sales Report from Retail Express" & vbNewLine & _
    "    " & Chr(176) & " Filters for 1 year ago to after last date in GP Report" & vbNewLine & _
    "    " & Chr(176) & " Do NOT Summarise by Outlet" & vbNewLine & _
    Chr(149) & " Copy all of the cells from the Gross Profit Report - Fulfilled.csv and add it to ""Sheet1"" of the Sales Report Export." & vbNewLine & _
    Chr(149) & " Sales Report Data should still be on ""Sheet2""." & vbNewLine & _
    vbNewLine & "All Set!" & vbNewLine & vbNewLine & _
    "If your Sheets are set up correctly, click Ok." & vbNewLine & "Otherwise Cancel to Abort.", _
    vbOKCancel + vbExclamation, "Slaes Report Fulilled Basis Macro Whole Year")
    'Removed this line and may add it back later if I find a way to deal with Quotes.
    'vbTab & Chr(176) & " Under Sales Status (Ctrl + Click) on Quote to add it." & vbNewLine & _

    
    If result = 2 Then
        Exit Sub
    End If
    
    'Test the Sheets
    If Not Worksheets("Sheet1").Cells(1, "R").Value = "Textbox3" Then
        MsgBox "Expected Sheet1 To contain Gross Profit Report - Fulfilled.csv data but it does not. This macro will now abort."
        Exit Sub
    End If

    
    If Not Worksheets("Sheet2").Cells(1, "E").Value = "OrderGuid" Then
        MsgBox "Wrong Data is in Sheet2. Expected Sales report data to be there but it was not. Please ensure that it is. This Macro will now abort"
        Exit Sub
    End If
    
'validation test finished
Application.ScreenUpdating = False
'Begin Sheet preparation

    

    'remove surplus columns
    Worksheets("Sheet2").Activate
    Columns("A:C").Select
    Selection.Delete Shift:=xlToLeft
    Columns("B:T").Select
    Selection.Delete Shift:=xlToLeft
    Worksheets("Sheet1").Activate
    Columns("A:B").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:G").Select
    Selection.Delete Shift:=xlToLeft
    Columns("F:G").Select
    Selection.Delete Shift:=xlToLeft
    Columns("K:O").Select
    Selection.Delete Shift:=xlToLeft
    
    
    'Create Needed Values
    Dim ws As Worksheet
    Dim LRS(104) As Long
    Dim LR(19) As Long

    'set the last row
    Worksheets("Sheet2").Activate
    LR(2) = Cells(Rows.count, "A").End(xlUp).row
    Worksheets("Sheet1").Activate
    LR(1) = Cells(Rows.count, "A").End(xlUp).row
    
    'Cross Reference Sales Person and Store
    Range("K1").Formula = "Sales Person"
    Range("L1").Formula = "Customer Type"
    Range("K2").Formula = "=Vlookup(B2, Sheet2!$A$1:$C$" & LR(2) & ", 2)"
    Range("L2").Formula = "=Vlookup(B2, Sheet2!$A$1:$C$" & LR(2) & ", 3)"
    Range("K2:L" & LR(1)).FillDown
    Range("K2:L" & LR(1)).Copy
    Range("K2").Select
    ActiveCell.PasteSpecial xlPasteValues
    Worksheets("Sheet2").Columns("A:C").Delete Shift:=xlToLeft
    
'End Sheet Preparation
'Select for Relevant Data

    'Filter for Fulfilled Sales
    
    Rows("1:1").Select
    Selection.AutoFilter
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort.SortFields.Add Key:=Range _
        ("E1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("Sheet1").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Fill out the Data
    Columns("F:F").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("I:I").Copy
    Range("F1").Select
    ActiveCell.PasteSpecial xlPasteAll
    Columns("I:I").Delete Shift:=xlToLeft
    Columns("G:G").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("G1").Formula = "Sales (Inc)"
    Range("F1").Formula = "SalesEx"
    Columns("H:H").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("H1").Formula = "Total Tax"
    Columns("I:I").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("K:K").Copy
    Range("I1").Select
    ActiveCell.PasteSpecial xlPasteAll
    Columns("J:K").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("O:P").Copy
    Range("J1").Select
    ActiveCell.PasteSpecial xlPasteAll
    Columns("J:J").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("O:O").Copy
    Range("J1").Select
    ActiveCell.PasteSpecial xlPasteAll
    Columns("N:Q").Delete Shift:=xlToLeft
    Columns("L:L").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("N:N").Copy
    Range("L1").Select
    ActiveCell.PasteSpecial xlPasteAll
    Columns("N:N").Delete Shift:=xlToLeft
    Range("G2").Formula = "=F2*1.1"
    Range("H2").Formula = "=F2*0.1"
    Range("G2:H" & LR(1)).FillDown
    
'End Select For relevant Data
'Start Week Logic
    
    Columns("E:E").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("E1").Formula = "WeekNumber"
    Columns("E:E").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("E1").Formula = "WeekDay"
    
    For i = 2 To LR(1)
        Cells(i, "F") = Format(Range("G" & i).Value, "ww", FirstDayOfWeek:=vbMonday, FirstWeekOfYear:=vbFirstFourDays)
        Cells(i, "E") = Format(Range("G" & i).Value, "w", FirstDayOfWeek:=vbMonday, FirstWeekOfYear:=vbFirstFourDays)
    Next
    
    Dim firstWeek As Integer
    Dim firstDay As Integer
    Dim firstdate As Date
    Dim lastdate As Date
    Dim lastWeek As Integer
    Dim lastDay As Integer
    firstWeek = Range("F2").Value
    firstDay = Range("E2").Value
    firstdate = Range("G2").Value
    lastdate = Range("G" & LR(1)).Value
    lastWeek = Range("F" & LR(1)).Value
    lastDay = Range("E" & LR(1)).Value
    
    If Not firstDay = 1 Then
        firstdate = firstdate + (8 - firstDay)
        firstWeek = firstWeek + 1
    End If
    If lastDay = 6 Then
        lastdate = lastdate + 1
        lastDay = lastDay + 1
    ElseIf Not lastDay = 7 Then
        lastdate = lastdate - lastDay
        lastDay = 7
        lastWeek = lastWeek - 1
    End If
    
    
    Dim iteration As Integer
    iteration = 2
    While firstdate < lastdate
        If iteration > 2 Then
        Set ws = Sheets.Add(After:=Sheets(Worksheets.count))
        End If
        
        Worksheets("Sheet1").Activate
        Worksheets("Sheet1").Range("A1:P" & LR(1)).AutoFilter _
        field:=7, _
        Criteria1:=">=" & firstdate, _
        Criteria2:="<=" & firstdate + 6, _
        VisibleDropDown:=False
        
        Range("A1:P" & LR(1)).Copy
        Worksheets("Sheet" & iteration).Activate
        Range("A1").Select
        ActiveCell.PasteSpecial xlPasteValues
        LRS(iteration) = Cells(Rows.count, "A").End(xlUp).row
        
        iteration = iteration + 1
        firstdate = firstdate + 7
        
        Worksheets("Sheet1").Activate
        ActiveSheet.ShowAllData
    Wend
    
    'Add Sheets
    Set ws = Sheets.Add(After:=Sheets(Worksheets.count))
    ws.Name = "Art Retail"  'LR(6)
    Set ws = Sheets.Add(After:=Sheets(Worksheets.count))
    ws.Name = "Art Prtnr"   'LR(7)
    Set ws = Sheets.Add(After:=Sheets(Worksheets.count))
    ws.Name = "Art WS"      'LR(8)
    Set ws = Sheets.Add(After:=Sheets(Worksheets.count))
    ws.Name = "Hab Retail"  'LR(9)
    Set ws = Sheets.Add(After:=Sheets(Worksheets.count))
    ws.Name = "Hab Prtnr"   'LR(10)
    Set ws = Sheets.Add(After:=Sheets(Worksheets.count))
    ws.Name = "Hab WS"      'LR(11)
    Set ws = Sheets.Add(After:=Sheets(Worksheets.count))
    ws.Name = "DIY"         'LR(12)
    Set ws = Sheets.Add(After:=Sheets(Worksheets.count))
    ws.Name = "T Art Retail"  'LR(13)
    ws.Activate
    Call weekHeaders
    Call salesHeaders
    Set ws = Sheets.Add(After:=Sheets(Worksheets.count))
    ws.Name = "T Art Prtnr"   'LR(14)
    ws.Activate
    Call weekHeaders
    Call salesHeaders
    Set ws = Sheets.Add(After:=Sheets(Worksheets.count))
    ws.Name = "T Art WS"      'LR(15)
    ws.Activate
    Call weekHeaders
    Call salesHeaders
    Set ws = Sheets.Add(After:=Sheets(Worksheets.count))
    ws.Name = "T Hab Retail"  'LR(16)
    ws.Activate
    Call weekHeaders
    Call salesHeaders
    Set ws = Sheets.Add(After:=Sheets(Worksheets.count))
    ws.Name = "T Hab Prtnr"   'LR(17)
    ws.Activate
    Call weekHeaders
    Call salesHeaders
    Set ws = Sheets.Add(After:=Sheets(Worksheets.count))
    ws.Name = "T Hab WS"      'LR(18)
    ws.Activate
    Call weekHeaders
    Call salesHeaders
    Set ws = Sheets.Add(After:=Sheets(Worksheets.count))
    ws.Name = "T DIY"         'LR(19)
    ws.Activate
    Call weekHeaders
    Call salesHeaders
    
'Start While loop to segment Weeks Data and Compile totals
    Dim sheetNumber As Integer
    sheetNumber = 2
    While sheetNumber < iteration
    
        'Simplify the Range
        Worksheets("Sheet" & sheetNumber).Activate
        Dim DatRange As Range
        Set DatRange = Range("A1:P" & LRS(sheetNumber))
        firstdate = DatRange.Range("G2").Value
        
        'Filter for DIY
            DatRange.AutoFilter field:=15, Criteria1:="Admin D", VisibleDropDown:=False
            DatRange.SpecialCells(xlCellTypeVisible).Copy
            Worksheets("DIY").Activate
            Range("A1").Activate
            ActiveCell.PasteSpecial (xlPasteAll)
            LR(12) = Cells(Rows.count, "A").End(xlUp).row
            Worksheets("Sheet" & sheetNumber).Activate
            Worksheets("Sheet" & sheetNumber).ShowAllData
            
            
        'Filter for Artarmon
            DatRange.AutoFilter field:=15, Criteria1:="<>Admin D", VisibleDropDown:=False
            DatRange.AutoFilter field:=1, Criteria1:="Artarmon", VisibleDropDown:=False
            DatRange.AutoFilter field:=16, Criteria1:="Retail", VisibleDropDown:=False
            DatRange.SpecialCells(xlCellTypeVisible).Copy
            Worksheets("Art Retail").Activate
            Range("A1").Activate
            ActiveCell.PasteSpecial (xlPasteAll)
            LR(6) = Cells(Rows.count, "A").End(xlUp).row
            
            Worksheets("Sheet" & sheetNumber).Activate
            DatRange.AutoFilter field:=16, Criteria1:="*Wholesale*", VisibleDropDown:=False
            DatRange.SpecialCells(xlCellTypeVisible).Copy
            Worksheets("Art WS").Activate
            Range("A1").Activate
            ActiveCell.PasteSpecial (xlPasteAll)
            LR(8) = Cells(Rows.count, "A").End(xlUp).row
            
            Worksheets("Sheet" & sheetNumber).Activate
            DatRange.AutoFilter field:=16, Criteria1:="*Partner*", Operator:=xlOr, Criteria2:="Employee", VisibleDropDown:=False
            DatRange.SpecialCells(xlCellTypeVisible).Copy
            Worksheets("Art Prtnr").Activate
            Range("A1").Activate
            ActiveCell.PasteSpecial (xlPasteAll)
            LR(7) = Cells(Rows.count, "A").End(xlUp).row
            
        'Filter for Haberfield
            Worksheets("Sheet" & sheetNumber).Activate
            DatRange.AutoFilter field:=1, Criteria1:="Haberfield", VisibleDropDown:=False
            DatRange.AutoFilter field:=16, Criteria1:="Retail", VisibleDropDown:=False
            DatRange.SpecialCells(xlCellTypeVisible).Copy
            Worksheets("Hab Retail").Activate
            Range("A1").Activate
            ActiveCell.PasteSpecial (xlPasteAll)
            LR(9) = Cells(Rows.count, "A").End(xlUp).row
            
            Worksheets("Sheet" & sheetNumber).Activate
            DatRange.AutoFilter field:=16, Criteria1:="*Wholesale*", VisibleDropDown:=False
            DatRange.SpecialCells(xlCellTypeVisible).Copy
            Worksheets("Hab WS").Activate
            Range("A1").Activate
            ActiveCell.PasteSpecial (xlPasteAll)
            LR(11) = Cells(Rows.count, "A").End(xlUp).row
            
            Worksheets("Sheet" & sheetNumber).Activate
            DatRange.AutoFilter field:=16, Criteria1:="*Partner*", Operator:=xlOr, Criteria2:="Employee", VisibleDropDown:=False
            DatRange.SpecialCells(xlCellTypeVisible).Copy
            Worksheets("Hab Prtnr").Activate
            Range("A1").Activate
            ActiveCell.PasteSpecial (xlPasteAll)
            LR(10) = Cells(Rows.count, "A").End(xlUp).row
            
                
        'Get them Sub Totals
        Worksheets("Art Retail").Activate
            Call subTotalSales
            Range("A2:N2").Copy
            Worksheets("T Art Retail").Activate
            Range("A" & sheetNumber).Activate
            Call InputWeekData(firstdate)
            ActiveCell.PasteSpecial xlPasteValuesAndNumberFormats
            Worksheets("Art Retail").Activate
            Rows("1:" & LR(6) + 2).Delete Shift:=xlUp
            
        Worksheets("Art Prtnr").Activate
            Call subTotalSales
            Range("A2:N2").Copy
            Worksheets("T Art Prtnr").Activate
            Range("A" & sheetNumber).Activate
            Call InputWeekData(firstdate)
            ActiveCell.PasteSpecial xlPasteValuesAndNumberFormats
            Worksheets("Art Prtnr").Activate
            Rows("1:" & LR(7) + 2).Delete Shift:=xlUp
            
        Worksheets("Art WS").Activate
            Call subTotalSales
            Range("A2:N2").Copy
            Worksheets("T Art WS").Activate
            Range("A" & sheetNumber).Activate
            Call InputWeekData(firstdate)
            ActiveCell.PasteSpecial xlPasteValuesAndNumberFormats
            Worksheets("Art WS").Activate
            Rows("1:" & LR(8) + 2).Delete Shift:=xlUp
            
        Worksheets("Hab Retail").Activate
            Call subTotalSales
            Range("A2:N2").Copy
            Worksheets("T Hab Retail").Activate
            Range("A" & sheetNumber).Activate
            Call InputWeekData(firstdate)
            ActiveCell.PasteSpecial xlPasteValuesAndNumberFormats
            Worksheets("Hab Retail").Activate
            Rows("1:" & LR(9) + 2).Delete Shift:=xlUp
            
        Worksheets("Hab Prtnr").Activate
            Call subTotalSales
            Range("A2:N2").Copy
            Worksheets("T Hab Prtnr").Activate
            Range("A" & sheetNumber).Activate
            Call InputWeekData(firstdate)
            ActiveCell.PasteSpecial xlPasteValuesAndNumberFormats
            Worksheets("Hab Prtnr").Activate
            Rows("1:" & LR(10) + 2).Delete Shift:=xlUp
            
        Worksheets("Hab WS").Activate
            Call subTotalSales
            Range("A2:N2").Copy
            Worksheets("T Hab WS").Activate
            Range("A" & sheetNumber).Activate
            Call InputWeekData(firstdate)
            ActiveCell.PasteSpecial xlPasteValuesAndNumberFormats
            Worksheets("Hab WS").Activate
            Rows("1:" & LR(11) + 2).Delete Shift:=xlUp
            
        Worksheets("DIY").Activate
            Call subTotalSales
            Range("A2:N2").Copy
            Worksheets("T DIY").Activate
            Range("A" & sheetNumber).Activate
            Call InputWeekData(firstdate)
            ActiveCell.PasteSpecial xlPasteValuesAndNumberFormats
            Worksheets("DIY").Activate
            Rows("1:" & LR(12) + 2).Delete Shift:=xlUp
            
        sheetNumber = sheetNumber + 1
    Wend ' While Loop then goes to the next week if there is one
    
'End While Loop to Segment the weeks Data
'Finish Gathering Data
'Make Sheet Pretty

    Application.ScreenUpdating = True
    
    ActiveWorkbook.Application.DisplayAlerts = False
        'Delete
        For i = 2 To iteration - 1
            Worksheets("Sheet" & i).Delete
        Next
        
        Worksheets("Art Retail").Delete
        Worksheets("Art Prtnr").Delete
        Worksheets("Art WS").Delete
        Worksheets("Hab Retail").Delete
        Worksheets("Hab Prtnr").Delete
        Worksheets("Hab WS").Delete
        Worksheets("DIY").Delete
        
        'Rename
        Worksheets("T Art Retail").Name = "Art Retail"
        Worksheets("T Art Prtnr").Name = "Art Partner"
        Worksheets("T Art WS").Name = "Art Wholesale"
        Worksheets("T Hab Retail").Name = "Hab Retail"
        Worksheets("T Hab Prtnr").Name = "Hab Partner"
        Worksheets("T Hab WS").Name = "Hab Wholesale"
        Worksheets("T DIY").Name = "DIY"
    
    ActiveWorkbook.Application.DisplayAlerts = True
    
    result = MsgBox("Your Sales Report has been filtered on a Fulfilment basis for " & startDate & " - " & endDate & " and is ready to be copied into the Google Sheet Sales Report.", vbOK + vbExclamation, "SalesReportFulfilledBasis")


End Sub

Sub subTotalSales()

    Dim LR As Integer
    LR = Cells(Rows.count, "A").End(xlUp).row + 2
    Rows("1:2").EntireRow.Insert Shift:=xlDown
    
    If LR = 3 Then
        Range("A1").Select
        Call salesHeaders
        Range("A2").Formula = "$0"
        Range("B2").Formula = "$0"
        Range("C2").Formula = "$0"
        Range("D2").Formula = "$0"
        Range("E2").Formula = "$0"
        Range("F2").Formula = "0%"
        Range("G2").Formula = "0"
        Range("H2").Formula = "$0"
        Range("I2").Formula = "0"
        Range("J2").Formula = "0"
        Range("K2").Formula = "$0"
        Range("L2").Formula = "$0"
        Range("M2").Formula = "0"
        Range("N2").Formula = "0%"
    Else
        Range("A1").Select
        Call salesHeaders
        With ActiveSheet
            .Cells(2, 1).Formula = "=Sum(H4:H" & LR & ")"   'Sales (Exc)
            .Cells(2, 2).Formula = "=Sum(I4:I" & LR & ")"   'Sales Inc
            .Cells(2, 3).Formula = "=Sum(J4:J" & LR & ")"   'total tax
            .Cells(2, 4).Formula = "=Sum(K4:K" & LR & ")"   'COGS
            .Cells(2, 5).Formula = "=Sum(L4:L" & LR & ")"   'GP Val
            .Cells(2, 6).Formula = "=E2/A2" '----------------GP%
            .Cells(2, 7).Formula = "=SUMPRODUCT(1/COUNTIF(B4:B" & LR & ",B4:B" & LR & "))"   'Transactions
            .Cells(2, 8).Formula = "=B2/G2" '----------------Avg Transaction Value
            .Cells(2, 9).Formula = "=Sum(N4:N" & LR & ")"   'Toatl Units
            .Cells(2, 10).Formula = "=I2/G2"    '------------Units/Transaction
            .Cells(2, 11).Formula = "=B2/I2"    '------------Average Unit Price
            .Cells(2, 12).Formula = "Not Calculated"    '----Discounts
            .Cells(2, 13).Formula = "Not Calculated"    '----No Discounted Transactions
            .Cells(2, 14).Formula = "Not calculated"    '----Discount %
            .Cells(2, 1).NumberFormat = "$###,##0.00"
            .Cells(2, 2).NumberFormat = "$###,##0.00"
            .Cells(2, 3).NumberFormat = "$###,##0.00"
            .Cells(2, 4).NumberFormat = "$###,##0.00"
            .Cells(2, 5).NumberFormat = "$###,##0.00"
            .Cells(2, 6).NumberFormat = "0.00%"
            .Cells(2, 7).NumberFormat = "0"
            .Cells(2, 8).NumberFormat = "$##0"
            .Cells(2, 9).NumberFormat = "0"
            .Cells(2, 10).NumberFormat = "0.00"
            .Cells(2, 11).NumberFormat = "$##0"
        End With
    End If

End Sub

Sub salesHeaders()
        ActiveCell.Formula = "Sales (Exc)"
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Formula = "Sales (Inc)"
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Formula = "Total Tax"
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Formula = "COGS (Exc)"
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Formula = "GP Value (Exc)"
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Formula = "GP %"
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Formula = "Transactions"
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Formula = "Avg Trans Value (inc)"
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Formula = "Total Units"
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Formula = "Units/Trans"
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Formula = "Avg Unit Price (inc)"
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Formula = "Discounts (inc)"
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Formula = "No. Disc. Trans"
        ActiveCell.Offset(0, 1).Select
        ActiveCell.Formula = "Discount %"
        ActiveCell.Offset(0, 1).Select
End Sub

Sub weekHeaders()

    ActiveCell.Formula = "Week Number"
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Formula = "Monday"
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Formula = "Sunday"
    ActiveCell.Offset(0, 1).Select
    
End Sub

Sub InputWeekData(x As Date)

    ActiveCell = Format(x, "ww", vbMonday, vbFirstFourDays)
    ActiveCell.Offset(0, 1).Select
    ActiveCell = x
    ActiveCell.Offset(0, 1).Select
    ActiveCell = x + 6
    ActiveCell.Offset(0, 1).Select

End Sub



