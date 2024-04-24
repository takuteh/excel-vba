Attribute VB_Name = "Module2"
Sub gotopop() '掲示表に情報を転記

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
slast_row = Sheets("集計表").Cells(Rows.Count, srank_col).End(xlUp).Row
slevel_col = search_slevel
n = 5
m = 1

'掲示表を消去
With Sheets("掲示表")
    .Range(.Cells(5, "A"), .Cells(40, "J")).Borders.LineStyle = xlLineStyleNone '罫線をクリア
    .Range(.Cells(5, "A"), .Cells(40, "J")).Clear 'セル内容クリア
End With

Application.ScreenUpdating = False '画面描画を停止

'集計表順位の最初から最後まで
For i = srank_row + 1 To slast_row - 1

    '0ポイントの場合飛ばす
    If Sheets("集計表").Cells(i, srank_col + 2) = 0 Then: GoTo CONTINUE
    
    With Sheets("掲示表")
        '掲示表セルの書式設定
        .Cells(n, 3 * m - 1).HorizontalAlignment = xlCenter
        .Cells(n, 3 * m).HorizontalAlignment = xlCenter
        .Cells(n, 3 * m - 1).VerticalAlignment = xlCenter
        .Cells(n, 3 * m).VerticalAlignment = xlCenter
        .Cells(n, 3 * m - 1).Font.Size = 18
        .Cells(n, 3 * m).Font.Size = 18
        .Cells(n, 3 * m - 1).Font.Bold = True
        .Cells(n, 3 * m).Font.Bold = True
        
        '掲示表に転記
        .Cells(n, 3 * m - 1) = Sheets("集計表").Cells(i, srank_col + 1) '名前コピー
        .Cells(n, 3 * m) = Sheets("集計表").Cells(i, srank_col + 2) 'ポイントコピー
        .Range(.Cells(n, 3 * m - 1), .Cells(n, 3 * m)).Borders.LineStyle = xlContinuous '罫線を引く
    End With
        
    n = n + 1 '行をインクリメント
    
CONTINUE:
    '集計表の順位が1位に戻ったら
    If Sheets("集計表").Cells(i + 1, srank_col) = 1 Then
       
        '現在の行と次の行の性別が違ったら列をリセット
        If Sheets("集計表").Cells(i, slevel_col + 1) <> Sheets("集計表").Cells(i + 1, slevel_col + 1) Then
            m = 0
            n = 25
        Else:
        
            '次の行が男性なら行を5にセット
             If Sheets("集計表").Cells(i + 1, slevel_col + 1) = "M" Then
                  n = 5
             ElseIf Sheets("集計表").Cells(i + 1, slevel_col + 1) = "F" Then
            '女性なら行を15にセット
                 n = 25
             End If
             
           
        End If
         m = m + 1 'レベル移行
    End If
    
    Next i
    

Call last_of_lebel
Application.ScreenUpdating = True '画面描画を実行

End Sub
Sub last_of_lebel() 'それぞれのレベルの最後の日付を取得&掲示表に更新日を入力

Dim sdate_row As Integer
Dim sdate_col As Integer
Dim i As Integer
Dim last_ad As String
Dim last_in As String
Dim last_ba As String

sdate_row = Range(search_sdate).Row
sdate_col = Range(search_sdate).Column
i = Sheets("集計表").Cells(sdate_row, Columns.Count).End(xlToLeft).Column

With Sheets("集計表")
    For i = 1 To Sheets("集計表").Cells(sdate_row, Columns.Count).End(xlToLeft).Column + 1
        If .Cells(sdate_row + 1, i) = "INAD" Then
            last_ad = .Cells(sdate_row, i)
            Sheets("掲示表").Cells(4, "C") = Format(last_ad, "mm/dd") + "更新"
            
            If last_in < .Cells(sdate_row, i) Then
                last_in = .Cells(sdate_row, i)
            End If
            Sheets("掲示表").Cells(4, "F") = Format(last_in, "mm/dd") + "更新"
            
        ElseIf .Cells(sdate_row + 1, i) = "BAIN" Then
            last_ba = .Cells(sdate_row, i)
            Sheets("掲示表").Cells(4, "I") = Format(last_ba, "mm/dd") + "更新"
            
             
            If last_in < .Cells(sdate_row, i) Then
                last_in = .Cells(sdate_row, i)
            End If
            Sheets("掲示表").Cells(4, "F") = Format(last_in, "mm/dd") + "更新"
            
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



'変数の初期化
ano_col = Range(search_ano(ActiveSheet.Name)).Column
first_row = Range(search_ano(ActiveSheet.Name)).Row + 1 'asの名前開始列
last_row = Sheets(ActiveSheet.Name).Cells(Rows.Count, ano_col + 1).End(xlUp).Row 'asの名前最後列
def_color = Sheets(ActiveSheet.Name).Cells(1, "A").Interior.Color '(1,1)のセル色
n = 1

Application.ScreenUpdating = False '画面描画を停止

'タイトル名変更
With Sheets("勝敗表")
.Shapes("textbox1").Delete
.Shapes.AddTextbox(1, 8, 30, 580, 60).Name = "textbox1"
.Shapes("textbox1").TextFrame.Characters.Text = "フレンドリーマッチ　" + s_level + " 勝敗表"
.Shapes("textbox1").Fill.Visible = msoFalse
.Shapes("textbox1").TextFrame2.TextRange.Font.NameFarEast = "HGP創英角ﾎﾟｯﾌﾟ体" '日本語
.Shapes("textbox1").TextFrame2.TextRange.Font.NameAscii = "HGP創英角ﾎﾟｯﾌﾟ体"   'ローマ字
.Shapes("textbox1").TextFrame.Characters.Font.Color = rgbOrangeRed
.Shapes("textbox1").TextFrame2.TextRange.Font.Size = 36
.Shapes("textbox1").Line.Visible = msoFalse
.Cells(7, "J") = s_date
End With

'勝敗表の名前欄を初期化
For i = 1 To 12
    Sheets("勝敗表").Cells(3 * i + 8, "B") = ""
        '勝敗表セルの書式設定
    With Sheets("勝敗表").Cells(3 * i + 8, "B")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Size = 18
        .Font.Name = "HG丸ｺﾞｼｯｸM-PRO"
    End With
Next i

For i = first_row To last_row

     'セルの色が白以外ならスキップ
    cel_color = Sheets(ActiveSheet.Name).Cells(i, ano_col + 1).Interior.Color
    If cel_color <> def_color Then: GoTo CONTINUE
    
    re_name = Replace(Sheets(ActiveSheet.Name).Cells(i, ano_col + 1), " ", "　")
    

    '勝敗表に名前入力
    cheak_length = Replace(Sheets(ActiveSheet.Name).Cells(i, ano_col + 1), " ", "　")
    Sheets("勝敗表").Cells(3 * n + 8, "B") = cheak_length
    If Len(cheak_length) > 5 Then
        Sheets("勝敗表").Cells(3 * n + 8, "B").Font.Size = 16
        Sheets("勝敗表").Cells(3 * n + 8, "B") = cheak_length
    End If
    
    If InStr(Sheets("勝敗表").Cells(3 * n + 8, "B"), "　") = 0 Then
        MsgBox ("苗字と名前の間にスペースを入れてください!")
        Exit Sub
    End If
    
    n = n + 1
    
CONTINUE:
Next i

'Call change_wllevel

Application.ScreenUpdating = True '画面描画を再開
MsgBox ("完了しました")
End Sub
