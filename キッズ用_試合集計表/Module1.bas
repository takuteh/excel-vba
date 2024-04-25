Attribute VB_Name = "Module1"
Sub レベル順ソート()
Attribute レベル順ソート.VB_ProcData.VB_Invoke_Func = " \n14"
'
' レベル順ソート Macro
'

'
    Range("A4:P154").Select
    ActiveWorkbook.Worksheets("集計表").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("集計表").Sort.SortFields.add Key:=Range("I5:I154"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("集計表").Sort.SortFields.add Key:=Range("H5:H154"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("集計表").Sort.SortFields.add Key:=Range("P5:P154"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("集計表").Sort
        .SetRange Range("A4:P154")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
