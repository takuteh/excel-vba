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

'������
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

'�`���~
Application.ScreenUpdating = False

For i = 5 To Cells(Rows.Count, "A").End(xlUp).Row + 100

    With Sheets("�����W�v�\")
        If next_month <> 6 Then
            '�O���܂ł̃|�C���g��l�݂̂ɂ���
            Range(Cells(i, jun), Cells(i, old_col)).Value = Range(Cells(i, jun), Cells(i, old_col)).Value
        End If
        
        '�Z���Ɋ֐����ߍ���
        .Cells(i, next_col).Formula = "=SUMPRODUCT(($K" & i & ":$O" & i & "=$T$5:$T$16)*($U$5:$U$16))"
    
        '���ʋL���������Z�b�g
        .Range(Cells(i, "K"), Cells(i, "O")).ClearContents
        
        '����6���̏ꍇ�͑S�ă��Z�b�g
        If next_month = 6 Then
            Range(Cells(i, aug), Cells(i, apr)).ClearContents
        End If
        
    End With
Next i

'�`��ĊJ
Application.ScreenUpdating = True
MsgBox ("�������܂���")



End Sub
