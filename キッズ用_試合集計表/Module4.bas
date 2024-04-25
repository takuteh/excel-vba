Attribute VB_Name = "Module4"
Sub delete()
Dim last_row As Integer

last_row = Cells(Rows.Count, 1).End(xlUp).Row
Range("A" & i, "Q" & i).delete
MsgBox (i & "s‚ğíœ‚µ‚Ü‚µ‚½")
If Cells(i, "Q") = "" Then

End If


Next i

MsgBox ("complete")
End Sub
