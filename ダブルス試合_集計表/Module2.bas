Attribute VB_Name = "Module2"
Sub gotopop() '�f���\�ɏ���]�L

Dim i As Integer

Dim cell_ad As Integer
Dim cell_in As Integer
Dim cell_ba As Integer
Dim srank_row As Integer
Dim srank_col As Integer
Dim slast_row As Integer
Dim slevel_col As Integer
Dim n As Integer
Dim m As Integer

srank_row = Range(search_srank).Row
srank_col = Range(search_srank).Column
slast_row = Sheets("�W�v�\").Cells(Rows.Count, srank_col).End(xlUp).Row
slevel_col = search_slevel
n = 5
m = 1

'�f���\������
With Sheets("�f���\")
    .Range(.Cells(5, "A"), .Cells(40, "J")).Borders.LineStyle = xlLineStyleNone '�r�����N���A
    .Range(.Cells(5, "A"), .Cells(40, "J")).Clear '�Z�����e�N���A
End With

Application.ScreenUpdating = False '��ʕ`����~

'�W�v�\���ʂ̍ŏ�����Ō�܂�
For i = srank_row + 1 To slast_row - 1

    '0�|�C���g�̏ꍇ��΂�
    If Sheets("�W�v�\").Cells(i, srank_col + 2) = 0 Then: GoTo CONTINUE
    
    With Sheets("�f���\")
        '�f���\�Z���̏����ݒ�
        .Cells(n, 3 * m - 1).HorizontalAlignment = xlCenter
        .Cells(n, 3 * m).HorizontalAlignment = xlCenter
        .Cells(n, 3 * m - 1).VerticalAlignment = xlCenter
        .Cells(n, 3 * m).VerticalAlignment = xlCenter
        .Cells(n, 3 * m - 1).Font.Size = 18
        .Cells(n, 3 * m).Font.Size = 18
        .Cells(n, 3 * m - 1).Font.Bold = True
        .Cells(n, 3 * m).Font.Bold = True
        
        '�f���\�ɓ]�L
        .Cells(n, 3 * m - 1) = Sheets("�W�v�\").Cells(i, srank_col + 1) '���O�R�s�[
        .Cells(n, 3 * m) = Sheets("�W�v�\").Cells(i, srank_col + 2) '�|�C���g�R�s�[
        .Range(.Cells(n, 3 * m - 1), .Cells(n, 3 * m)).Borders.LineStyle = xlContinuous '�r��������
    End With
        
    n = n + 1 '�s���C���N�������g
    
CONTINUE:
    '�W�v�\�̏��ʂ�1�ʂɖ߂�����
    If Sheets("�W�v�\").Cells(i + 1, srank_col) = 1 Then
       
        '���݂̍s�Ǝ��̍s�̐��ʂ�������������Z�b�g
        If Sheets("�W�v�\").Cells(i, slevel_col + 1) <> Sheets("�W�v�\").Cells(i + 1, slevel_col + 1) Then
            m = 0
            n = 25
        Else:
        
            '���̍s���j���Ȃ�s��5�ɃZ�b�g
             If Sheets("�W�v�\").Cells(i + 1, slevel_col + 1) = "M" Then
                  n = 5
             ElseIf Sheets("�W�v�\").Cells(i + 1, slevel_col + 1) = "F" Then
            '�����Ȃ�s��15�ɃZ�b�g
                 n = 25
             End If
             
           
        End If
         m = m + 1 '���x���ڍs
    End If
    
    Next i
    

Call last_of_lebel
Application.ScreenUpdating = True '��ʕ`������s

End Sub
Sub last_of_lebel() '���ꂼ��̃��x���̍Ō�̓��t���擾&�f���\�ɍX�V�������

Dim sdate_row As Integer
Dim sdate_col As Integer
Dim i As Integer
Dim last_ad As String
Dim last_in As String
Dim last_ba As String

sdate_row = Range(search_sdate).Row
sdate_col = Range(search_sdate).Column
i = Sheets("�W�v�\").Cells(sdate_row, Columns.Count).End(xlToLeft).Column

With Sheets("�W�v�\")
    For i = 1 To Sheets("�W�v�\").Cells(sdate_row, Columns.Count).End(xlToLeft).Column + 1
        If .Cells(sdate_row + 1, i) = "INAD" Then
            last_ad = .Cells(sdate_row, i)
            Sheets("�f���\").Cells(4, "C") = Format(last_ad, "mm/dd") + "�X�V"
            
            If last_in < .Cells(sdate_row, i) Then
                last_in = .Cells(sdate_row, i)
            End If
            Sheets("�f���\").Cells(4, "F") = Format(last_in, "mm/dd") + "�X�V"
            
        ElseIf .Cells(sdate_row + 1, i) = "BAIN" Then
            last_ba = .Cells(sdate_row, i)
            Sheets("�f���\").Cells(4, "I") = Format(last_ba, "mm/dd") + "�X�V"
            
             
            If last_in < .Cells(sdate_row, i) Then
                last_in = .Cells(sdate_row, i)
            End If
            Sheets("�f���\").Cells(4, "F") = Format(last_in, "mm/dd") + "�X�V"
            
        End If
        
    Next i
End With

End Sub
Sub gotowinlose()

Dim i As Integer
Dim ano_col As Integer
Dim first_row As Integer
Dim last_row As Integer
Dim n As Integer
Dim def_color As Long
Dim cel_color As Long
Dim rename As String
Dim s_cnt As Integer
Dim cheak_length As String

Call Module1.take_DaLe



'�ϐ��̏�����
ano_col = Range(search_ano(ActiveSheet.Name)).Column
first_row = Range(search_ano(ActiveSheet.Name)).Row + 1 'as�̖��O�J�n��
last_row = Sheets(ActiveSheet.Name).Cells(Rows.Count, ano_col + 1).End(xlUp).Row 'as�̖��O�Ō��
def_color = Sheets(ActiveSheet.Name).Cells(1, "A").Interior.Color '(1,1)�̃Z���F
n = 1

Application.ScreenUpdating = False '��ʕ`����~

'�^�C�g�����ύX
With Sheets("���s�\")
.Shapes("textbox1").Delete
.Shapes.AddTextbox(1, 8, 30, 580, 60).Name = "textbox1"
.Shapes("textbox1").TextFrame.Characters.Text = "�t�����h���[�}�b�`�@" + s_level + " ���s�\"
.Shapes("textbox1").Fill.Visible = msoFalse
.Shapes("textbox1").TextFrame2.TextRange.Font.NameFarEast = "HGP�n�p�p�߯�ߑ�" '���{��
.Shapes("textbox1").TextFrame2.TextRange.Font.NameAscii = "HGP�n�p�p�߯�ߑ�"   '���[�}��
.Shapes("textbox1").TextFrame.Characters.Font.Color = rgbOrangeRed
.Shapes("textbox1").TextFrame2.TextRange.Font.Size = 36
.Shapes("textbox1").Line.Visible = msoFalse
.Cells(7, "J") = s_date
End With

'���s�\�̖��O����������
For i = 1 To 12
    Sheets("���s�\").Cells(3 * i + 8, "B") = ""
        '���s�\�Z���̏����ݒ�
    With Sheets("���s�\").Cells(3 * i + 8, "B")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 18
        .Font.Name = "HG�ۺ޼��M-PRO"
    End With
Next i

For i = first_row To last_row

     '�Z���̐F�����ȊO�Ȃ�X�L�b�v
    cel_color = Sheets(ActiveSheet.Name).Cells(i, ano_col + 1).Interior.Color
    If cel_color <> def_color Then: GoTo CONTINUE
    
    re_name = Replace(Sheets(ActiveSheet.Name).Cells(i, ano_col + 1), " ", "�@")
    

    '���s�\�ɖ��O����
    cheak_length = Replace(Sheets(ActiveSheet.Name).Cells(i, ano_col + 1), " ", "�@")
    Sheets("���s�\").Cells(3 * n + 8, "B") = cheak_length
    If Len(cheak_length) > 5 Then
        Sheets("���s�\").Cells(3 * n + 8, "B").Font.Size = 16
        Sheets("���s�\").Cells(3 * n + 8, "B") = cheak_length
    End If
    
    If InStr(Sheets("���s�\").Cells(3 * n + 8, "B"), "�@") = 0 Then
        MsgBox ("�c���Ɩ��O�̊ԂɃX�y�[�X�����Ă�������!")
        Exit Sub
    End If
    
    n = n + 1
    
CONTINUE:
Next i

'Call change_wllevel

Application.ScreenUpdating = True '��ʕ`����ĊJ
MsgBox ("�������܂���")
End Sub
