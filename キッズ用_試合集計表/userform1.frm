VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "登録"
   ClientHeight    =   2415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3045
   OleObjectBlob   =   "userform1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "userform1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
x = Cells(Rows.Count, 1).End(xlUp).Row + 1

If TextBox1 = "" Or TextBox2 = "" Or TextBox3 = "" Then

MsgBox ("未入力の項目があります｡")

Else

For i = 1 To x
If Cells(i + 4, "A") = TextBox1 Then
MsgBox ("既にあります。")
End
End If
Next i

Cells(x, "A") = TextBox1
Cells(x, "P") = TextBox2
Cells(x, "J").Value = TextBox3.Value

MsgBox (x & "行に追加しました。")

Unload Me
Cells(x, "K").Select
End If


End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub


