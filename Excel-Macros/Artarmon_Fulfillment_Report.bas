Attribute VB_Name = "Module11"
Sub Artarmon_Fulfillment_Report_Export_XLS()
Attribute Artarmon_Fulfillment_Report_Export_XLS.VB_ProcData.VB_Invoke_Func = "k\n14"

'
' Fulfillment_Report_Export_XLS Macro
'
' Keyboard Shortcut: Ctrl+k
'
    Application.ScreenUpdating = False
    
    Range( _
        "A:A,F:F,H:O,S:V,Y:Z,AC:AD,AF:AK,AM:AM" _
        ).Select
    Range("AG1").Activate
    Selection.Delete Shift:=xlToLeft
   
   
    Cells.Select
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range _
        ("I2:I10000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range _
        ("A2:A10000"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.ActiveSheet.Sort
        .SetRange Range("A1:AK10000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With


    Columns("C:E").Select
    Selection.Cut
    Range("O1").Select
    ActiveSheet.Paste
    Columns("C:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("J:J").Select
    Selection.Insert Shift:=xlToRight
    
    

col = "a"
    LR = Cells(Rows.count, col).End(xlUp).row
    fr = 1
    tr = 1
    
    Application.DisplayAlerts = False
    
    Do Until fr >= LR
        Do While Cells(tr + 1, col) = Cells(fr, col)
            tr = tr + 1
        Loop
        Range(Cells(fr, col), Cells(tr, col)).Merge
        fr = tr + 1
        tr = fr
    Loop
    
    Application.DisplayAlerts = True
    
Range("A1").Select
    ActiveCell.Offset(1).Select
    Selection.Copy
    
    Do Until IsEmpty(ActiveCell)
         
         ActiveCell.Offset(0, 1).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        ActiveCell.Offset(0, 1).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        ActiveCell.Offset(0, 1).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        ActiveCell.Offset(0, 1).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        ActiveCell.Offset(0, 1).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        ActiveCell.Offset(0, 1).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        ActiveCell.Offset(0, 1).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        ActiveCell.Offset(0, 1).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
        ActiveCell.Offset(1, -8).Select
        Selection.Copy
        
    Loop
    
   
Columns("F:F").Select
    Selection.NumberFormat = "[$-F800]dddd, mmmm dd, yyyy"
    
    

col = "F"
    LR = Cells(Rows.count, col).End(xlUp).row
    fr = 1
    tr = 1
    
    Application.DisplayAlerts = False
    
    Do Until fr >= LR
        Do While Cells(tr + 1, col) = Cells(fr, col)
            tr = tr + 1
        Loop
        Range(Cells(fr, col), Cells(tr, col)).Merge
        fr = tr + 1
        tr = fr
    Loop
    
    Application.DisplayAlerts = True

   Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("G:G").Select
    Selection.Cut
    Columns("A:A").Select
    ActiveSheet.Paste
    
    Columns("C:C").Select
    Selection.Cut
    Columns("Q:Q").Select
    ActiveSheet.Paste
    
    Columns("C:C").Select
    Selection.Delete Shift:=xlToLeft
   
    
    Columns("F:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
    
    Columns("D:D").Select
    Selection.Cut
    Columns("F:F").Select
    ActiveSheet.Paste
    
    Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft
   
    Columns("J:J").Select
    Selection.Cut
    Columns("F:F").Select
    ActiveSheet.Paste
   
    Columns("I:I").Select
    Selection.Cut
    Columns("G:G").Select
    ActiveSheet.Paste
    
    Columns("H:H").Select
    Selection.Cut
    Columns("R:R").Select
    ActiveSheet.Paste
   
    Columns("Q:Q").Select
    Selection.Cut
    Columns("H:H").Select
    ActiveSheet.Paste
    
    Columns("N:P").Select
    Selection.Cut
    Columns("I:I").Select
    ActiveSheet.Paste
    
    Columns("N:Q").Select
    Selection.Delete Shift:=xlToLeft
      
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "ID"

    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Qty"
 
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "$"
        
    Columns("H:H").Select
    Selection.Replace What:="Awaiting Payment", Replacement:="AP", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="Processed", Replacement:="P", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Columns("D:E").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    Columns("D:E").EntireColumn.AutoFit
    
    'Invoice Status
    Columns("H:H").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    Selection.ColumnWidth = 3
    
    'ID Column
    Columns("I:I").Select
    Selection.ColumnWidth = 6.5
    
    'Manfucature SKU Column
    Columns("J:J").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.ColumnWidth = 12
    
    'Qty Column
    Columns("L:L").Select
    Selection.ColumnWidth = 4
    With Selection
        .HorizontalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    'Delmode Column
    Range("N1").Select
    ActiveCell.FormulaR1C1 = "DelMode"
    Columns("N:N").Select
    Selection.ColumnWidth = 8
    With Selection
        .HorizontalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With

    'DelDue Column
    Columns("A:A").Select
    With Selection
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    
    Application.PrintCommunication = True
    Application.ScreenUpdating = True
    
    ActiveWorkbook.Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:="P:\Public Folder\ArtarmonFulfilmentReport.htm", _
        FileFormat:=xlHtml, ReadOnlyRecommended:=False, CreateBackup:=False, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
    ActiveWorkbook.Application.DisplayAlerts = True
    
End Sub

