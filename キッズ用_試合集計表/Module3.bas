Attribute VB_Name = "Module3"
Sub level_sort()
Attribute level_sort.VB_ProcData.VB_Invoke_Func = " \n14"
'
' level_sort Macro
'

'
    Range("A5:P201").Select
    ActiveWorkbook.Worksheets("総合集計表").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("総合集計表").Sort.SortFields.add Key:=Range("I5:I201") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("総合集計表").Sort.SortFields.add Key:=Range("H5:H201") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("総合集計表").Sort.SortFields.add Key:=Range("P5:P201") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("総合集計表").Sort
        .SetRange Range("A4:P" & Cells(Rows.Count, 1).End(xlUp).Row)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub name_sort()
Attribute name_sort.VB_ProcData.VB_Invoke_Func = " \n14"
'
' name_sort Macro
'

'
    Range("A5:P200").Select
    ActiveWorkbook.Worksheets("総合集計表").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("総合集計表").Sort.SortFields.add Key:=Range("P5:P200") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("総合集計表").Sort
        .SetRange Range("A4:P" & Cells(Rows.Count, 1).End(xlUp).Row)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
