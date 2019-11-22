Sub prospectsInfluencers()
'
' prospectsInfluencers Macro
'

'
     
    Sheets("Advanced Search Prospect Export").Select
    Range("A1").Select
    Sheets("Advanced Search Prospect Export").Select
    
    Selection.AutoFilter
    
    ActiveSheet.Range("$A$1:$CG$100000").AutoFilter Field:=80, Criteria1:="=No", _
        Operator:=xlOr, Criteria2:="="
   
    
    ActiveSheet.Range("$A:$CG").AutoFilter Field:=18, Criteria1:="<>"
    ActiveCell.Offset(0, 17).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Range("A1").Select
    ActiveCell.Select
    ActiveSheet.Paste
    Sheets("Advanced Search Prospect Export").Select
    
    ActiveSheet.Range("$A$1:$CG$100000").AutoFilter Field:=18
    ActiveSheet.Range("$A$1:$CG$100000").AutoFilter Field:=19, Criteria1:="<>"
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Sheet1").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveSheet.Paste
    Sheets("Advanced Search Prospect Export").Select
    ActiveSheet.Range("$A$1:$CG$100000").AutoFilter Field:=19
    ActiveSheet.Range("$A$1:$CG$100000").AutoFilter Field:=20
    
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveSheet.Range("$A$1:$CG$100000").AutoFilter Field:=19
    ActiveSheet.Range("$A$1:$CG$100000").AutoFilter Field:=20, Criteria1:="<>"
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Sheet1").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveSheet.Paste
    Sheets("Advanced Search Prospect Export").Select
    ActiveSheet.Range("$A$1:$CG$100000").AutoFilter Field:=20
    ActiveSheet.Range("$A$1:$CG$100000").AutoFilter Field:=50, Criteria1:="<>"
    ActiveCell.Offset(0, 30).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Sheet1").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveSheet.Paste
    Sheets("Advanced Search Prospect Export").Select
    
    ActiveWindow.LargeScroll Down:=-1
    ActiveSheet.Range("$A$1:$CG$100000").AutoFilter Field:=50
    ActiveSheet.Range("$A$1:$CG$100000").AutoFilter Field:=51, Criteria1:="<>"
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Sheet1").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(0, 1).Range("A1").Select
    Sheets("Advanced Search Prospect Export").Select
    ActiveSheet.Range("$A$1:$CG$100000").AutoFilter Field:=51
    
    ActiveSheet.Range("$A$1:$CG$100000").AutoFilter Field:=52, Criteria1:="<>"
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Sheet1").Select
    ActiveCell.Select
    ActiveSheet.Paste
    Sheets("Advanced Search Prospect Export").Select
    ActiveSheet.Range("$A$1:$CG$100000").AutoFilter Field:=52
    
    ActiveSheet.Range("$A$1:$CG$100000").AutoFilter Field:=66, Criteria1:="<>"
    ActiveCell.Offset(0, 14).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Sheet1").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveSheet.Paste
    Sheets("Advanced Search Prospect Export").Select
    ActiveSheet.Range("$A$1:$CG$100000").AutoFilter Field:=66
    
    ActiveSheet.Range("$A$1:$CG$100000").AutoFilter Field:=67, Criteria1:="<>"
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Sheet1").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveSheet.Paste
    Sheets("Advanced Search Prospect Export").Select
    ActiveSheet.Range("$A$1:$CG$100000").AutoFilter Field:=67
    Sheets("Advanced Search Prospect Export").Select
    ActiveSheet.Range("$A$1:$CG$100000").AutoFilter Field:=67
    
    ActiveSheet.Range("$A$1:$CG$100000").AutoFilter Field:=68, Criteria1:="<>"
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Sheet1").Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveSheet.Paste
    
    
    Sheets("Sheet1").Select
    ActiveCell.Rows("1:1").EntireRow.Select
    Selection.Delete Shift:=xlUp
    Range("B1").Select
    
    

    Range("B1:B100000").Select
    Application.CutCopyMode = False
    Selection.Cut
    ActiveCell.Offset(1, -1).Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select

    
    ActiveCell.Offset(0, 2).Range("A1").Select
    Range("C1:C100000").Select
    Selection.Cut
    ActiveCell.Offset(0, -2).Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select
    

    ActiveCell.Offset(0, 3).Range("A1").Select
    Range("D1:D100000").Select
    Selection.Cut
    ActiveCell.Offset(0, -3).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
   ActiveSheet.Paste
    Selection.End(xlUp).Select



    ActiveCell.Offset(0, 4).Range("A1").Select
    Range("E1:E100000").Select
    Selection.Cut
    ActiveCell.Offset(0, -4).Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    Selection.End(xlUp).Select

    
    ActiveCell.Offset(0, 5).Range("A1").Select
    Range("F1:F100000").Select
    Selection.Cut
    ActiveCell.Offset(0, -5).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveCell.Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(-1, 0).Range("A1").Select
    Selection.End(xlUp).Select
    
    ActiveCell.Offset(0, 6).Range("A1").Select
    Range("G1:G100000").Select
    Selection.Cut
    ActiveCell.Offset(0, -6).Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    Range("A1").Select

    ActiveCell.Offset(0, 7).Range("A1").Select
    Range("H1:H100000").Select
    Selection.Cut
    ActiveCell.Offset(0, -7).Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    Range("A1").Select
    
        ActiveCell.Offset(0, 8).Range("A1").Select
    Range("I1:I100000").Select
    Selection.Cut
    ActiveCell.Offset(0, -8).Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    Range("A1").Select
    
    Sheets("Advanced Search Prospect Export").Select
    ActiveSheet.ShowAllData
    Selection.AutoFilter
    Sheets("Sheet1").Select
End Sub