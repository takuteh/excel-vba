Attribute VB_Name = "Module4"
Public return_f3 As Integer
'
'Sub register() '—\–ñ“ú“o˜^
'Dim do_row As Integer
'Dim user_row As Integer
'Dim user_col As Integer
'Dim redate As Integer
'Dim as_name As String
'
'as_name = ActiveSheet.Name
'do_row = ActiveCell.Row
'user_row = Range(search_user).Row + 1
'user_col = Range(search_user).Column
'redate = Range(search_redate).Column
'
'If Cells(do_row, redate) <> "" Or Cells(do_row, redate + 1) <> "" Then
'    return_f3 = 1
'    UserForm3.Show
'Else:
'  do_input_of_register
'End If
'
'
'
'End Sub
'Sub money() '“ü‹à“ú“o˜^
'Dim do_row As Integer
'Dim user_row As Integer
'Dim user_col As Integer
'Dim mndate As Integer
'Dim as_name As String
'
'as_name = ActiveSheet.Name
'do_row = ActiveCell.Row
'user_row = Range(search_user).Row + 1
'user_col = Range(search_user).Column
'mndate = Range(search_mndate).Column
'
'If Cells(do_row, mndate) <> "" Or Cells(do_row, mndate + 1) <> "" Then
'      return_f3 = 2
'      UserForm3.Show
'Else:
'  do_input_of_money
'End If
'
'End Sub
'Function do_input_of_register()
'Dim do_row As Integer
'Dim user_row As Integer
'Dim user_col As Integer
'Dim redate As Integer
'Dim as_name As String
'
'as_name = ActiveSheet.Name
'do_row = ActiveCell.Row
'user_row = Range(search_user).Row + 1
'user_col = Range(search_user).Column
'redate = Range(search_redate).Column
'    Cells(do_row, redate) = Date
'    Cells(do_row, redate + 1) = Sheets(ActiveSheet.Name).Cells(user_row, user_col)
'End Function
'Function do_input_of_money()
'Dim do_row As Integer
'Dim user_row As Integer
'Dim user_col As Integer
'Dim mndate As Integer
'Dim as_name As String
'
'as_name = ActiveSheet.Name
'do_row = ActiveCell.Row
'user_row = Range(search_user).Row + 1
'user_col = Range(search_user).Column
'mndate = Range(search_mndate).Column
'    Cells(do_row, mndate) = Date
'    Cells(do_row, mndate + 1) = Sheets(ActiveSheet.Name).Cells(user_row, user_col)
'End Function
'
'Function search_user() As String 'g—pÒ‚ÌˆÊ’u‚ğ’T‚·
'
'Dim c As Range
'Set c = Sheets(ActiveSheet.Name).Range("A1:H15").Find("g—pÒ«")
'search_user = c.Address(False, False)
'
'End Function
'Function search_redate() As String '—\–ñ“ú‚ÌˆÊ’u‚ğ’T‚·
'
'Dim c As Range
'Set c = Sheets(ActiveSheet.Name).Range("A1:H15").Find("—\–ñ“ú")
'search_redate = c.Address(False, False)
'
'End Function
'Function search_mndate() As String '“ü‹à“ú‚ÌˆÊ’u‚ğ’T‚·
'
'Dim c As Range
'Set c = Sheets(ActiveSheet.Name).Range("A1:H15").Find("“ü‹à“ú")
'search_mndate = c.Address(False, False)
'
'End Function
