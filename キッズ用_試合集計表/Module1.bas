Attribute VB_Name = "Module1"
Sub ���x�����\�[�g()
Attribute ���x�����\�[�g.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ���x�����\�[�g Macro
'

'
    Range("A4:P154").Select
    ActiveWorkbook.Worksheets("�W�v�\").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�W�v�\").Sort.SortFields.add Key:=Range("I5:I154"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("�W�v�\").Sort.SortFields.add Key:=Range("H5:H154"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("�W�v�\").Sort.SortFields.add Key:=Range("P5:P154"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�W�v�\").Sort
        .SetRange Range("A4:P154")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
