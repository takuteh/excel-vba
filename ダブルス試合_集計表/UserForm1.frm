VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "�o�^"
   ClientHeight    =   2490
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2925
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Dim ssl As Integer
Dim ssra_col As Integer

ssra_col = Range(search_srank).Column '�W�v�\�̏��ʗ��̗�
ssl = search_slevel

If ComboBox1 = "BA" Then
    Sheets("�W�v�\").Cells(s_cnt, ssl) = "3BA"
ElseIf ComboBox1 = "IN" Then
    Sheets("�W�v�\").Cells(s_cnt, ssl) = "2IN"
ElseIf ComboBox1 = "AD" Then
    Sheets("�W�v�\").Cells(s_cnt, ssl) = "1AD"
End If

If ComboBox2 = "�j��" Then
    Sheets("�W�v�\").Cells(s_cnt, ssl + 1) = "M"
ElseIf ComboBox2 = "����" Then
    Sheets("�W�v�\").Cells(s_cnt, ssl + 1) = "F"
End If

If ComboBox1 = "" Or ComboBox2 = "" Then
    MsgBox ("���͘R�ꂪ����܂�")
Else:
    Sheets("�W�v�\").Cells(s_cnt, ssra_col + 1) = p_name
    Unload Me
End If

End Sub

Private Sub CommandButton2_Click()
MsgBox ("�����𒆒f���܂�")
return_form = 1
Unload Me
End Sub



Private Sub UserForm_Initialize()
Label5 = p_name

With ComboBox1
.AddItem "BA"
.AddItem "IN"
.AddItem "AD"
End With


With ComboBox2
.AddItem "�j��"
.AddItem "����"
End With

End Sub
