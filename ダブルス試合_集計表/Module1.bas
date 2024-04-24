Attribute VB_Name = "Module1"
Public p_name As String
Public s_cnt As Integer
Public return_form As Integer
Public s_level As String
Public s_date As String
Sub singles()

Dim ssl As Integer
Dim cnt As Integer
Dim date_last As Integer
Dim date_cnt As Integer
Dim as_name As String
Dim def_color As Long
Dim cel_color As Long
Dim delete_last As Integer
Dim sano_row As Integer
Dim sano_col As Integer
Dim sara As Integer
Dim ssda As Integer
Dim ssra_row As Integer
Dim ssra_col As Integer


'�ϐ��̏�����
return_form = 0
as_name = ActiveSheet.Name      '�A�N�e�B�u�V�[�g�����擾
sano_row = Range(search_ano(as_name)).Row
sano_col = Range(search_ano(as_name)).Column
sara = Range(search_arank(as_name)).Column
ssda = Range(search_sdate).Row         '�W�v�\�̓��t���̍s
ssra_row = Range(search_srank).Row     '�W�v�\�̏��ʗ��̍s
ssra_col = Range(search_srank).Column  '�W�v�\�̏��ʗ��̗�
def_color = Sheets(as_name).Cells(1, "A").Interior.Color '(1,1)�̃Z���F
date_last = Sheets("�W�v�\").Cells(ssda, Columns.Count).End(xlToLeft).Column  '�W�v�\�̓��t���̍Ō��
s_last_column = Sheets("�W�v�\").Cells(Rows.Count, ssra_col + 1).End(xlUp).Row '�W�v�\�̍ŏI�s


Call take_DaLe

'�W�v�\�ɓ������t�����邩
For date_cnt = 1 To date_last
    '���t��str�^�ɒ����Ĕ�r
    If CStr(Sheets("�W�v�\").Cells(ssda, date_cnt)) = Format(s_date, "yyyy/mm/dd") Then
        Exit For
    End If
Next date_cnt

'�������t���Ȃ��ꍇ�ŏI�s�ɒǉ�
If date_cnt = date_last + 1 Then
     '�ŏI�s�ǉ��A���t����
    With Sheets("�W�v�\")
        .Columns(date_cnt).Insert
        .Columns(date_cnt).ClearFormats
        .Cells(ssda, date_cnt) = s_date
    End With
Else:
    '�������t������ꍇ��x���ʂ�����
    delete_last = Sheets("�W�v�\").Cells(Rows.Count, date_cnt).End(xlUp).Row
    Range(Sheets("�W�v�\").Cells(ssra_row + 1, date_cnt), Sheets("�W�v�\").Cells(delete_last, date_cnt)).ClearContents
End If
'���x������
Sheets("�W�v�\").Cells(ssda + 1, date_cnt) = s_level

'No1~���O�����͂���Ă���Ō�̗��܂ŌJ��Ԃ�
For cnt = sano_row + 1 To Sheets(as_name).Cells(Rows.Count, sano_col + 1).End(xlUp).Row
    '�Z���̐F�����ȊO�Ȃ�X�L�b�v
    cel_color = Sheets(as_name).Cells(cnt, sano_col + 1).Interior.Color
    If cel_color <> def_color Then: GoTo CONTINUE

    
    p_name = Replace(Sheets(as_name).Cells(cnt, sano_col + 1), " ", "�@") '���p�X�y�[�X->�S�p
    
    If InStr(p_name, "�@") = 0 Then
        MsgBox ("�c���Ɩ��O�̊ԂɃX�y�[�X�����Ă�������!")
        Exit Sub
    End If
    


    
    '�W�v�\�ɓ������O�����邩
    For s_cnt = 1 To Sheets("�W�v�\").Cells(Rows.Count, ssra_col + 1).End(xlUp).Row
        If Sheets("�W�v�\").Cells(s_cnt, ssra_col + 1) = p_name Then: Exit For
    Next
 
    ssl = search_slevel '�W�v�\�̔F�苉�̗����
    
    '�������O�������ꍇ
    If s_cnt = Sheets("�W�v�\").Cells(Rows.Count, ssra_col + 1).End(xlUp).Row + 1 Or Sheets("�W�v�\").Cells(s_cnt, ssl) = "" Or Sheets("�W�v�\").Cells(s_cnt, ssl + 1) = "" Then
        '���[�U�[�t�H�[���ɏ�����͂�����
       UserForm1.Show
    
    End If
    
    '���[�U�[�t�H�[���̕Ԃ�l1�̎����[�v����
    If return_form = 1 Then: Exit For
    
    '�W�v�\�ɏ��ʂ����
    If s_level = "BAIN" Then
        Sheets("�W�v�\").Cells(s_cnt, date_cnt) = Sheets(as_name).Cells(cnt, sara) + 200
    ElseIf s_level = "INAD" Then
        Sheets("�W�v�\").Cells(s_cnt, date_cnt) = Sheets(as_name).Cells(cnt, sara) + 100
    End If
    
CONTINUE:
    Next cnt
    


If return_form = 0 Then
    Call sort_of_point
    MsgBox ("�������܂���")
End If

End Sub
Sub take_DaLe() '�V�[�g��������t���x���擾
Dim a As Integer
Dim check_text As String
s_level = ""
s_date = ""

For a = 1 To Len(ActiveSheet.Name)
    check_text = Mid(ActiveSheet.Name, a, 1)
    If check_text Like "[A-Z]" Then
        s_level = s_level & check_text '�V�[�g�����烌�x���𒊏o
    Else: s_date = s_date & check_text '���t�𒊏o
    End If
Next a
End Sub
Function search_ano(ByVal as_name As String) As String 'asNO�̈ʒu��T��
'No1�̃Z����T��
Dim c As Range
Set c = Sheets(as_name).Range("A1:J10").Find("NO")
search_ano = c.Address(False, False)

End Function
Function search_arank(ByVal as_name As String) As String 'as���ʂ̈ʒu��T��

Dim c As Range
Set c = Sheets(as_name).Range("A1:Z15").Find("����", searchorder:=xlByRows)
search_arank = c.Address(False, False)

End Function
Function search_srank() As String '�W�v�\���ʂ̈ʒu��T��

Dim c As Range
Set c = Sheets("�W�v�\").Range("A1:H15").Find("����")
search_srank = c.Address(False, False)

End Function
Public Function search_slevel() As Integer '�W�v�\�̔F�苉�̍s��T��

Dim i As Integer
i = 1
Do
    If Sheets("�W�v�\").Cells(Range(search_srank).Row, i) = "�F�苉" Then: Exit Do
    i = i + 1
Loop

search_slevel = i

End Function

Function search_sdate() As String '�W�v�\���t�̈ʒu��T��

Dim c As Range
Set c = Sheets("�W�v�\").Range("A1:Z15").Find("���t��")
search_sdate = c.Address(False, False)

End Function

