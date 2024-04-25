VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   1335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2340
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
ComboBox1.AddItem "6"
ComboBox1.AddItem "8"
ComboBox1.AddItem "10"
ComboBox1.AddItem "12"
ComboBox1.AddItem "2"
ComboBox1.AddItem "4"

End Sub

Private Sub CommandButton1_Click()
form_result = True
next_month = ComboBox1.Text
Unload Me
End Sub

Private Sub CommandButton2_Click()
form_result = False
Unload Me
End Sub
