Attribute VB_Name = "Module2"
Sub name_lookup()
Dim get_string As String '�{�^���̕������i�[

name_sort

'������̌���

With Sheets("�����W�v�\")



    get_string = .Shapes(Application.Caller).TextFrame.Characters.Text
    
    Select Case get_string
        Case "��"
            set_firstrow ("��")
        Case "��"
            set_firstrow ("��")
        Case "��"
            set_firstrow ("��")
        Case "��"
            set_firstrow ("��")
        Case "��"
            set_firstrow ("��")
        Case "��"
            set_firstrow ("��")
        Case "��"
            set_firstrow ("��")
        Case "��"
            set_firstrow ("��")
        Case "��"
            set_firstrow ("��")
        Case "��"
            set_firstrow ("��")
        Case Else
            MsgBox ("�s���ȑ���ł�")
    End Select
End With

End Sub
Function search_lname() As String '�c���̈ʒu��T��(last_name)

Dim c As Range
Set c = Sheets("�����W�v�\").Range("A1:U15").Find("�݂傤��")
search_lname = c.Address(False, False)

End Function
Function set_firstrow(lf_string As String) As Integer '�c���̓������̐擪�s��I��(lf_string->lastname_first_string)

Dim i As Integer
Dim lname_row As Integer '�c���̍s
Dim lname_column As Integer '�c���̗�

With Sheets("�����W�v�\")
     '������
    lname_column = .Range(search_lname).Column
    lname_row = .Range(search_lname).Row
    
    For i = 1 To .Cells(Rows.Count, 1).End(xlUp).Row
        If lf_string = Left(Cells(i, .Range(search_lname).Column), 1) Then
             .Cells(i, lname_column - 5).Select
                
                '�I�����ꂽ�Z�����ŏ�s�ɃX�N���[��
                With ActiveWindow
                    .ScrollRow = ActiveCell.Row
                End With
            
            Exit For
        End If
    Next i
    
    If i - 1 = .Cells(Rows.Count, 1).End(xlUp).Row Then
        MsgBox (lf_string + "�s�͂���܂���")
    End If

End With

End Function
