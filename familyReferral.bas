Sub familyReferral()
'
' Family Referral Macro
'

'
    Cells.Select
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
    ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range("B2:B2887") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.ActiveSheet.Sort
        .SetRange Range("A1:E2887")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$E$2887").AutoFilter Field:=3, Criteria1:="Family"
    ActiveSheet.Range("$A$1:$E$2887").AutoFilter Field:=4, Criteria1:="<>"
    ActiveSheet.Range("$A$1:$E$2887").AutoFilter Field:=5, Criteria1:= _
        "Current resident"
End Sub