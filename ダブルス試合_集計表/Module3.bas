Attribute VB_Name = "Module3"
Sub sort_of_date()
'
' sort_of_date Macro
'
Dim name_col
Dim date_row
Dim date_col
Dim last_col
Dim last_row

name_col = Range(search_srank).Column + 1
date_row = Range(search_sdate).Row
date_col = Range(search_sdate).Column + 1
last_col = Sheets("集計表").Cells(date_row, Columns.Count).End(xlToLeft).Column
last_row = Sheets("集計表").Cells(Rows.Count, name_col).End(xlUp).Row
'
    ActiveWorkbook.Worksheets("集計表").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("集計表").Sort.SortFields.Add Key:=Range(Cells(date_row, date_col), Cells(date_row, last_col)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("集計表").Sort
        .SetRange Range(Sheets("集計表").Cells(date_row, date_col), Sheets("集計表").Cells(last_row, last_col))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlLeftToRight
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub sort_of_point()
'
' sort_of_point Macro
'
Dim name_col
Dim name_row

name_col = Range(search_srank).Column + 1
name_row = Range(search_srank).Row + 1
last_col = search_slevel + 1
last_row = Sheets("集計表").Cells(Rows.Count, name_col).End(xlUp).Row


    ActiveWorkbook.Worksheets("集計表").Sort.SortFields.Clear
      ActiveWorkbook.Worksheets("集計表").Sort.SortFields.Add Key:=Range(Cells(name_row, last_col), Cells(last_row, last_col)), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("集計表").Sort.SortFields.Add Key:=Range(Cells(name_row, search_slevel), Cells(last_row, search_slevel)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
  
    ActiveWorkbook.Worksheets("集計表").Sort.SortFields.Add Key:=Range(Cells(name_row, name_col + 1), Cells(last_row, name_col + 1)), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("集計表").Sort
        .SetRange Range(Cells(name_row, name_col), Cells(last_row, last_col))
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Call sort_of_date
    Call rank_s
    Call gotopop
    Call last_of_lebel
    
End Sub


Sub rank_s() '順位の入力(S)

Dim rank_row As Integer
Dim rank_col As Integer
Dim sex_col As Integer
Dim name_col As Integer
Dim last_row As Integer
Dim i As Integer
Dim m As Integer

rank_row = Range(search_srank).Row
rank_col = Range(search_srank).Column
sex_col = search_slevel + 1
name_col = Range(search_srank).Column + 1
last_row = Sheets("集計表").Cells(Rows.Count, name_col).End(xlUp).Row
i = rank_row + 1
m = 1

For i = rank_row + 1 To last_row
    Sheets("集計表").Cells(i, rank_col) = m
    
    m = m + 1
    If Sheets("集計表").Cells(i, search_slevel) <> Sheets("集計表").Cells(i + 1, search_slevel) Then
        m = 1
    End If

Next i

End Sub

Sub rank_d() '順位の入力(D)

Dim rank_row As Integer
Dim rank_col As Integer
Dim sex_col As Integer
Dim name_col As Integer
Dim last_row As Integer
Dim i As Integer
Dim m As Integer

rank_row = Range(search_srank).Row
rank_col = Range(search_srank).Column
sex_col = search_slevel + 1
name_col = Range(search_srank).Column + 1
last_row = Sheets("集計表").Cells(Rows.Count, name_col).End(xlUp).Row
i = rank_row + 1
m = 1

For i = rank_row + 1 To last_row
    Sheets("集計表").Cells(i, rank_col) = m
    
    m = m + 1
    If Sheets("集計表").Cells(i, sex_col) <> Sheets("集計表").Cells(i + 1, sex_col) Then
        m = 1
    End If

Next i

End Sub
