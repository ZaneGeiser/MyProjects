Attribute VB_Name = "Module6"
Sub TradePriceGroupsUpdate()
Attribute TradePriceGroupsUpdate.VB_ProcData.VB_Invoke_Func = " \n14"
'
' TradePriceGroupsUpdate Macro
' In Price Group Mass Download (Fixed) from Retail Express, this macro will determine the margins on all products and then assign them into 4 groups for trade discount puropses. Group 1 : no discount :: group 2 : Low volume :: Group 3 : Core Products :: Group 4 : HR
'

'
'Remove Obsolete Discount Groups
    Columns("AA:AK").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft

'Inserting a Column at Column A
    Range("A1").EntireColumn.Insert
    
'Setting Margin Title
    Range("A10").Select
    ActiveCell.FormulaR1C1 = "Margin %"
    
'Filling Margin % values
    Range("A11").Select
    Selection.Style = "Percent"
    ActiveCell.Formula = "=((O11/1.1)-L11)/(O11/1.1)"
    'ActiveCell.FormulaR1C1 = "=((RC[14]/1.1)-RC[11])/(RC[14]/1.1)"
    Range("A11:A35000").Select
    Selection.FillDown
    Selection.Copy
    Range("A11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
'Inserting a Column at Column A
    Range("A1").EntireColumn.Insert
    
'Set Group Title
    Range("A10").Select
    ActiveCell.FormulaR1C1 = "Discount Group"
    
'Setting Groups based on Margin
    Range("A11").Select
    Selection.NumberFormat = "General"
    ActiveCell.Formula = "=IF(OR(B11<=0.15,C11=133780,C11=133782,C11=133253,C11=140406,C11=146046,C11=146089),""Group 1"",IF(B11<=0.25,""Group 2"",IF(AND(B11<1,OR(AND(G11=""Ce"", I11=""Rola"" ),AND(G11<>""Ce"", I11<>""Hayman Reese""))),""Group 3"",IF(AND(B11<1,OR(G11=""Ce"", I11=""Hayman Reese"")),""Group 4"",""Group 1""))))"
    Range("A11:A35000").Select
    Range("A35000").Activate
    Selection.FillDown
    Selection.Copy
    Range("A11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
'Removing invalid groups
    Range("C11").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, -2).Select
    ActiveCell.EntireRow.Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A11").Select
    Cells.Replace What:="#DIV/0!", Replacement:="Group 1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
'Defining Formulas for all discounts
    'Trade 1
    Range("U11").Formula = "=IF(A11=""Group 1"",""CLEAR DATA"",IF(A11=""Group 2"",P11*0.95,IF(A11=""Group 3"",P11*0.95,IF(A11=""Group 4"",P11*0.95,""CLEAR DATA""))))"
    'Trade 2
    Range("V11").Formula = "=IF(A11=""Group 1"",""CLEAR DATA"",IF(A11=""Group 2"",P11*0.95,IF(A11=""Group 3"",P11*0.90,IF(A11=""Group 4"",P11*0.90,""CLEAR DATA""))))"
    'Dealer 1
    Range("W11").Formula = "=IF(A11=""Group 1"",""CLEAR DATA"",IF(A11=""Group 2"",P11*0.95,IF(A11=""Group 3"",P11*0.90,IF(A11=""Group 4"",P11*0.85,""CLEAR DATA""))))"
    'Trade 1
    Range("X11").Formula = "=IF(A11=""Group 1"",""CLEAR DATA"",IF(A11=""Group 2"",P11*0.90,IF(A11=""Group 3"",P11*0.85,IF(A11=""Group 4"",P11*0.80,""CLEAR DATA""))))"
    'Trade 2
    Range("Y11").Formula = "=IF(A11=""Group 1"",""CLEAR DATA"",IF(A11=""Group 2"",P11*0.90,IF(A11=""Group 3"",P11*0.85,IF(A11=""Group 4"",P11*0.75,""CLEAR DATA""))))"
    'Trade 3
    Range("Z11").Formula = "=IF(A11=""Group 1"",""CLEAR DATA"",IF(A11=""Group 2"",P11*0.85,IF(A11=""Group 3"",P11*0.83,IF(A11=""Group 4"",P11*0.73,""CLEAR DATA""))))"
    'Trade 4
    Range("AA11").Formula = "=IF(A11=""Group 1"",""CLEAR DATA"",IF(A11=""Group 2"",P11*0.85,IF(A11=""Group 3"",P11*0.80,IF(A11=""Group 4"",P11*0.71,""CLEAR DATA""))))"
    'Trade 5
    Range("AB11").Formula = "=IF(A11=""Group 1"",""CLEAR DATA"",IF(A11=""Group 2"",P11*0.85,IF(A11=""Group 3"",P11*0.80,IF(A11=""Group 4"",P11*0.69,""CLEAR DATA""))))"
    
'Fill Discounts down
    Range("U11:AB35000").Select
    Selection.FillDown
    Selection.Copy
    Range("U11").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("U11:AB35000").NumberFormat = "\$#,##0.00;[Red]\$-#,##0.00"
    Range("M11:S35000").NumberFormat = "\$#,##0.00;[Red]\$-#,##0.00"
    
'Removing excess formulas
'    Range("U11:U35000").Select
'    ActiveWorkbook.Worksheets("Fixed Price Groups").Sort.SortFields.Clear
'    ActiveWorkbook.Worksheets("Fixed Price Groups").Sort.SortFields.Add Key:= _
'        Range("U11"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
'        xlSortTextAsNumbers
'    With ActiveWorkbook.Worksheets("Fixed Price Groups").Sort
'        .SetRange Range("A11:AB35000")
'        .Header = xlNo
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'
'    'Loop one step by 10000
'    Do Until ActiveCell = ""
'    ActiveCell.Offset(10000, 0).Select
'    Loop
'
'    'Loop 2 Step by 1000
'    Do Until ActiveCell <> ""
'    ActiveCell.Offset(-1000, 0).Select
'    Loop
'
'    'Loop 3 Step by 100
'    Do Until ActiveCell = ""
'    ActiveCell.Offset(100, 0).Select
'    Loop
'
'    'Loop 4 Step by 10
'        Do Until ActiveCell <> ""
'    ActiveCell.Offset(-10, 0).Select
'    Loop
'
'    'Loop 5 Step by 1
'    Do Until ActiveCell = ""
'    ActiveCell.Offset(1, 0).Select
'    Loop
    
    Range("C11").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, -2).Select
    ActiveCell.EntireRow.Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    
    
'    ActiveCell.EntireRow.Select
'    Range(Selection, Selection.End(xlDown)).Select
'    Selection.Delete Shift:=xlUp
'    Range(Selection, Selection.End(xlDown)).Select
'    Range(Selection, Selection.End(xlDown)).Select
'    Selection.Delete Shift:=xlUp
'    Range(Selection, Selection.End(xlDown)).Select
'    Selection.Delete Shift:=xlUp

    
'Removing Margin% Column A&B to prepare for upload to ReEx
    Columns("A:B").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    
    
'Adjusting Colums to make pretty
    Columns("A:A").ColumnWidth = 9
    Columns("B:B").ColumnWidth = 16
    Columns("C:C").ColumnWidth = 12
    Columns("D:D").ColumnWidth = 38
    Columns("E:E").ColumnWidth = 12
    Columns("F:F").ColumnWidth = 21
    Columns("G:J").ColumnWidth = 8
    Columns("K:O").ColumnWidth = 11
    Columns("P:P").ColumnWidth = 8
    Columns("Q:R").ColumnWidth = 3
    Columns("S:Z").ColumnWidth = 10
    
'Insert Instructions
    Range("D1").Select
    ActiveCell.FormulaR1C1 = _
        "PLEASE UPLOAD TO RETAIL EXPRESS UNDER SETTINGS > PRICE GROUPS > PRICE GROUPS MASS UPLOAD (FIXED)"
    With Selection
        .HorizontalAlignment = xlLeft
        .Orientation = 0
        .AddIndent = True
        .IndentLevel = 1
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
   End With
    Selection.Font.Bold = True
    With Selection.Font
        .Name = "Arial"
        .Size = 18
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With

'save
    ActiveWorkbook.Save
    
End Sub
