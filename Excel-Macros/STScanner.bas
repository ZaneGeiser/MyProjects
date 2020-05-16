Attribute VB_Name = "Module17"
Sub STScaner()
Attribute STScaner.VB_ProcData.VB_Invoke_Func = "s\n14"

Dim match As Range
Dim kitNumber As String
Dim Scanning As Boolean
Scanning = True

While Scanning
    Dim sku As Variant
    sku = InputBox("Please Scan IT!", "Scan to Add to your Count", "1234")
    
    If sku = "" Or sku = 1234 Then
        Beep
        Scanning = False
        Exit Sub
    End If
    
        LR = Cells(Rows.count, "A").End(xlUp).row
        'Change the column here to look for different data.
        sku = "K" & sku & "W"
        Set match = Worksheets(1).Range("C6:C" & LR).Find(sku, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=xlTrue)
        If match Is Nothing Then
            Beep
            Scanning = False
            Exit Sub
        Else
            match.Activate
            ActiveCell.Select
            ActiveCell.Offset(0, 7).Select 'Change this offset if you change the Column
            ActiveCell = ActiveCell.Value + 1
        End If

Wend
End Sub
