Attribute VB_Name = "Module2"
Sub name_lookup()
Dim get_string As String 'ボタンの文字を格納

name_sort

'文字列の検索

With Sheets("総合集計表")



    get_string = .Shapes(Application.Caller).TextFrame.Characters.Text
    
    Select Case get_string
        Case "あ"
            set_firstrow ("あ")
        Case "か"
            set_firstrow ("か")
        Case "さ"
            set_firstrow ("さ")
        Case "た"
            set_firstrow ("た")
        Case "な"
            set_firstrow ("な")
        Case "は"
            set_firstrow ("は")
        Case "ま"
            set_firstrow ("ま")
        Case "や"
            set_firstrow ("や")
        Case "ら"
            set_firstrow ("ら")
        Case "わ"
            set_firstrow ("わ")
        Case Else
            MsgBox ("不正な操作です")
    End Select
End With

End Sub
Function search_lname() As String '苗字の位置を探す(last_name)

Dim c As Range
Set c = Sheets("総合集計表").Range("A1:U15").Find("みょうじ")
search_lname = c.Address(False, False)

End Function
Function set_firstrow(lf_string As String) As Integer '苗字の頭文字の先頭行を選択(lf_string->lastname_first_string)

Dim i As Integer
Dim lname_row As Integer '苗字の行
Dim lname_column As Integer '苗字の列

With Sheets("総合集計表")
     '初期化
    lname_column = .Range(search_lname).Column
    lname_row = .Range(search_lname).Row
    
    For i = 1 To .Cells(Rows.Count, 1).End(xlUp).Row
        If lf_string = Left(Cells(i, .Range(search_lname).Column), 1) Then
             .Cells(i, lname_column - 5).Select
                
                '選択されたセルを最上行にスクロール
                With ActiveWindow
                    .ScrollRow = ActiveCell.Row
                End With
            
            Exit For
        End If
    Next i
    
    If i - 1 = .Cells(Rows.Count, 1).End(xlUp).Row Then
        MsgBox (lf_string + "行はありません")
    End If

End With

End Function
