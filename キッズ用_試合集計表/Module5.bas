Attribute VB_Name = "Module5"
Public next_month As Integer
Public form_result As Boolean


Sub UpdateMonth()

Dim i As Long
Dim old_col As Integer
Dim next_col As Integer
Dim aug As Integer
Dim oct As Integer
Dim des As Integer
Dim feb As Integer
Dim apr As Integer
Dim jun As Integer

'初期化
aug = 3
oct = 4
des = 5
feb = 6
apr = 7
jun = 2


UserForm2.Show

Select Case next_month
    Case 8
        next_col = aug
        old_col = jun
    Case 10
        next_col = oct
        old_col = aug
    Case 12
        next_col = des
        old_col = oct
    Case 2
        next_col = feb
        old_col = des
    Case 4
        next_col = apr
        old_col = feb
    Case 6
        UserForm3.Show
        
        next_col = jun
        old_col = apr
End Select
        
If form_result = False Then
   End
End If

'描画停止
Application.ScreenUpdating = False

For i = 5 To Cells(Rows.Count, "A").End(xlUp).Row + 100

    With Sheets("総合集計表")
        If next_month <> 6 Then
            '前月までのポイントを値のみにする
            Range(Cells(i, jun), Cells(i, old_col)).Value = Range(Cells(i, jun), Cells(i, old_col)).Value
        End If
        
        'セルに関数埋め込み
        .Cells(i, next_col).Formula = "=SUMPRODUCT(($K" & i & ":$O" & i & "=$T$5:$T$16)*($U$5:$U$16))"
    
        '順位記入欄をリセット
        .Range(Cells(i, "K"), Cells(i, "O")).ClearContents
        
        '次が6月の場合は全てリセット
        If next_month = 6 Then
            Range(Cells(i, aug), Cells(i, apr)).ClearContents
        End If
        
    End With
Next i

'描画再開
Application.ScreenUpdating = True
MsgBox ("完了しました")



End Sub
