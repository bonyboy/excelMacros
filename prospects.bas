Sub prospects()
'
' prospects Macro
'

'
    Range("A1:CG1892").Select
    ActiveWorkbook.Worksheets("Advanced Search Prospect Export").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Advanced Search Prospect Export").Sort.SortFields. _
        Add Key:=Range("CB2:CB100000"), SortOn:=xlSortOnValues, Order:=xlDescending _
        , DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Advanced Search Prospect Export").Sort
        .SetRange Range("A1:CG100000")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$CG$100000").AutoFilter Field:=80, Criteria1:="=No", _
        Operator:=xlOr, Criteria2:="="
    ActiveSheet.Range("$A$1:$CG$100000").AutoFilter Field:=18, Criteria1:="<>"
    Columns("R:R").Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:A").EntireColumn.AutoFit
    Sheets("Advanced Search Prospect Export").Select
    ActiveSheet.Range("$A$1:$CG$100000").AutoFilter Field:=18
    ActiveSheet.Range("$A$1:$CG$100000").AutoFilter Field:=19, Criteria1:="<>"
    Columns("S:S").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Sheet1").Select
    Range("B1").Select
    ActiveSheet.Paste
    Sheets("Advanced Search Prospect Export").Select
    ActiveSheet.Range("$A$1:$CG$100000").AutoFilter Field:=19
    ActiveSheet.Range("$A$1:$CG$100000").AutoFilter Field:=20, Criteria1:="<>"
    Columns("T:T").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Sheet1").Select
    Columns("C:C").Select
    ActiveSheet.Paste
    Sheets("Advanced Search Prospect Export").Select
    ActiveSheet.Range("$A$1:$CG$100000").AutoFilter Field:=20
    Application.CutCopyMode = False
    ActiveSheet.ShowAllData
    Selection.AutoFilter
    Range("A1").Select
    Sheets("Sheet1").Select
    Columns("B:D").Select
    Selection.ColumnWidth = 8.57
    Columns("B:D").EntireColumn.AutoFit
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("D1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveWindow.LargeScroll Down:=23
    Range("D:D").Find("").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("C2:C4").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("D:D").Find("").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:C").Select
    Range("C901").Activate
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    ActiveWindow.ScrollRow = 1
    Columns("A:C").EntireColumn.AutoFit
End Sub