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


'変数の初期化
return_form = 0
as_name = ActiveSheet.Name      'アクティブシート名を取得
sano_row = Range(search_ano(as_name)).Row
sano_col = Range(search_ano(as_name)).Column
sara = Range(search_arank(as_name)).Column
ssda = Range(search_sdate).Row         '集計表の日付欄の行
ssra_row = Range(search_srank).Row     '集計表の順位欄の行
ssra_col = Range(search_srank).Column  '集計表の順位欄の列
def_color = Sheets(as_name).Cells(1, "A").Interior.Color '(1,1)のセル色
date_last = Sheets("集計表").Cells(ssda, Columns.Count).End(xlToLeft).Column  '集計表の日付欄の最後尾
s_last_column = Sheets("集計表").Cells(Rows.Count, ssra_col + 1).End(xlUp).Row '集計表の最終行


Call take_DaLe

'集計表に同じ日付があるか
For date_cnt = 1 To date_last
    '日付をstr型に直して比較
    If CStr(Sheets("集計表").Cells(ssda, date_cnt)) = Format(s_date, "yyyy/mm/dd") Then
        Exit For
    End If
Next date_cnt

'同じ日付がない場合最終行に追加
If date_cnt = date_last + 1 Then
     '最終行追加、日付入力
    With Sheets("集計表")
        .Columns(date_cnt).Insert
        .Columns(date_cnt).ClearFormats
        .Cells(ssda, date_cnt) = s_date
    End With
Else:
    '同じ日付がある場合一度順位を消去
    delete_last = Sheets("集計表").Cells(Rows.Count, date_cnt).End(xlUp).Row
    Range(Sheets("集計表").Cells(ssra_row + 1, date_cnt), Sheets("集計表").Cells(delete_last, date_cnt)).ClearContents
End If
'レベル入力
Sheets("集計表").Cells(ssda + 1, date_cnt) = s_level

'No1~名前が入力されている最後の欄まで繰り返し
For cnt = sano_row + 1 To Sheets(as_name).Cells(Rows.Count, sano_col + 1).End(xlUp).Row
    'セルの色が白以外ならスキップ
    cel_color = Sheets(as_name).Cells(cnt, sano_col + 1).Interior.Color
    If cel_color <> def_color Then: GoTo CONTINUE

    
    p_name = Replace(Sheets(as_name).Cells(cnt, sano_col + 1), " ", "　") '半角スペース->全角
    
    If InStr(p_name, "　") = 0 Then
        MsgBox ("苗字と名前の間にスペースを入れてください!")
        Exit Sub
    End If
    


    
    '集計表に同じ名前があるか
    For s_cnt = 1 To Sheets("集計表").Cells(Rows.Count, ssra_col + 1).End(xlUp).Row
        If Sheets("集計表").Cells(s_cnt, ssra_col + 1) = p_name Then: Exit For
    Next
 
    ssl = search_slevel '集計表の認定級の列を代入
    
    '同じ名前が無い場合
    If s_cnt = Sheets("集計表").Cells(Rows.Count, ssra_col + 1).End(xlUp).Row + 1 Or Sheets("集計表").Cells(s_cnt, ssl) = "" Or Sheets("集計表").Cells(s_cnt, ssl + 1) = "" Then
        'ユーザーフォームに情報を入力させる
       UserForm1.Show
    
    End If
    
    'ユーザーフォームの返り値1の時ループ抜け
    If return_form = 1 Then: Exit For
    
    '集計表に順位を入力
    If s_level = "BAIN" Then
        Sheets("集計表").Cells(s_cnt, date_cnt) = Sheets(as_name).Cells(cnt, sara) + 200
    ElseIf s_level = "INAD" Then
        Sheets("集計表").Cells(s_cnt, date_cnt) = Sheets(as_name).Cells(cnt, sara) + 100
    End If
    
CONTINUE:
    Next cnt
    


If return_form = 0 Then
    Call sort_of_point
    MsgBox ("完了しました")
End If

End Sub
Sub take_DaLe() 'シート名から日付レベル取得
Dim a As Integer
Dim check_text As String
s_level = ""
s_date = ""

For a = 1 To Len(ActiveSheet.Name)
    check_text = Mid(ActiveSheet.Name, a, 1)
    If check_text Like "[A-Z]" Then
        s_level = s_level & check_text 'シート名からレベルを抽出
    Else: s_date = s_date & check_text '日付を抽出
    End If
Next a
End Sub
Function search_ano(ByVal as_name As String) As String 'asNOの位置を探す
'No1のセルを探す
Dim c As Range
Set c = Sheets(as_name).Range("A1:J10").Find("NO")
search_ano = c.Address(False, False)

End Function
Function search_arank(ByVal as_name As String) As String 'as順位の位置を探す

Dim c As Range
Set c = Sheets(as_name).Range("A1:Z15").Find("順位", searchorder:=xlByRows)
search_arank = c.Address(False, False)

End Function
Function search_srank() As String '集計表順位の位置を探す

Dim c As Range
Set c = Sheets("集計表").Range("A1:H15").Find("順位")
search_srank = c.Address(False, False)

End Function
Public Function search_slevel() As Integer '集計表の認定級の行を探す

Dim i As Integer
i = 1
Do
    If Sheets("集計表").Cells(Range(search_srank).Row, i) = "認定級" Then: Exit Do
    i = i + 1
Loop

search_slevel = i

End Function

Function search_sdate() As String '集計表日付の位置を探す

Dim c As Range
Set c = Sheets("集計表").Range("A1:Z15").Find("日付→")
search_sdate = c.Address(False, False)

End Function

