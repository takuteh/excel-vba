VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "UserForm3"
   ClientHeight    =   1305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3450
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ToggleButton1_Click()
If return_f3 = 1 Then
    do_input_of_register
ElseIf return_f3 = 2 Then
    do_input_of_money
End If

Unload Me
End Sub

Private Sub ToggleButton2_Click()
return_f3 = 0
Unload Me
End Sub

