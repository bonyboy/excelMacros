Sub InfluencerEmailCollect()
'
' InfluencerEmailCollect Macro
'
    Range("AX1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$CI$84").AutoFilter Field:=50, Criteria1:="<>"
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Advanced Search Prospect Export").Select
    Application.CutCopyMode = False
    ActiveSheet.ShowAllData
    Range("AY1").Select
    ActiveSheet.Range("$A$1:$CI$84").AutoFilter Field:=51, Criteria1:="<>"
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Sheet1").Select
    Range("B1").Select
    ActiveSheet.Paste
    Sheets("Advanced Search Prospect Export").Select
    Application.CutCopyMode = False
    ActiveSheet.ShowAllData
    ActiveSheet.Range("$A$1:$CI$84").AutoFilter Field:=52, Criteria1:="<>**", _
        Operator:=xlAnd
    Range("AZ1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Sheet1").Select
    Range("C1").Select
    ActiveSheet.Paste
    Sheets("Advanced Search Prospect Export").Select
    Application.CutCopyMode = False
    ActiveSheet.ShowAllData
    ActiveSheet.Range("$A$1:$CI$84").AutoFilter Field:=66, Criteria1:="<>"
    Range("BN1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Sheet1").Select
    Range("D1").Select
    ActiveSheet.Paste
    Sheets("Advanced Search Prospect Export").Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    Range("BO1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$CI$84").AutoFilter Field:=67, Criteria1:="<>"
    Range("BO1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Sheet1").Select
    Range("E1").Select
    ActiveSheet.Paste
    Sheets("Advanced Search Prospect Export").Select
    Application.CutCopyMode = False
    ActiveSheet.ShowAllData
    Columns("BP:BP").Select
    ActiveSheet.Range("$A$1:$CI$84").AutoFilter Field:=68, Criteria1:="<>**", _
        Operator:=xlAnd
    Range("BP1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Sheet1").Select
    Range("F1").Select
    ActiveSheet.Paste
    Sheets("Advanced Search Prospect Export").Select
    Application.CutCopyMode = False
    ActiveSheet.ShowAllData
    Selection.AutoFilter
    Sheets("Sheet1").Select

    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Range("A1").Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Range("B1:B100000").Select
    Selection.Cut
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select

    
    ActiveCell.Offset(0, 3).Range("A1").Select
    Range("C1:C100000").Select
    Selection.Cut
    ActiveCell.Offset(0, -2).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveCell.Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(-1, 0).Range("A1").Select
    Selection.End(xlUp).Select
    
    ActiveCell.Offset(0, 4).Range("A1").Select
    Range("D1:D100000").Select
    Selection.Cut
    ActiveCell.Offset(0, -3).Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    Range("A1").Select

    ActiveCell.Offset(0, 5).Range("A1").Select
    Range("E1:E100000").Select
    Selection.Cut
    ActiveCell.Offset(0, -4).Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    Range("A1").Select
    
        ActiveCell.Offset(0, 6).Range("A1").Select
    Range("F1:F100000").Select
    Selection.Cut
    ActiveCell.Offset(0, -5).Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    Range("A1").Select
    
End Sub