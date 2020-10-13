Dim i As Integer
Dim j As Integer


Sub dup()

For i = 4 To 127

If Cells(i, 4).Value = "F" Then

Cells(i, 24).Value = Cells(i, 3).Value


End If


Next i

If Cells(i, 4).Value <> "F" Then

For i = 4 To 127

For j = i + 9 To 127

If Cells(i, 2).Text = Cells(j, 2).Text Then

If Cells(i, 4).Text = "F" Then
Cells(j, 8).Value = ""
Else

Cells(j, 26).Value = Cells(j, 3).Value

End If

End If

Next j

Next i

End If

End Sub

Sub cls()

For i = 4 To 127

Cells(i, 24).Value = ""
Cells(i, 26).Value = ""
Next i

End Sub
