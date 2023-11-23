Attribute VB_Name = "Module1"
Public Sub TIMER()
'If Form5.year.Caption = form2.year.Caption Then
'form2.Timer1.Interval = 300 And Form5.Timer1.Interval = 300
'year = year + 10
'End If

End Sub

Public Sub CHECKED()
If Form7.Option5.Value = True Then
Form7.Text1 = form2.Label3.Caption
End If

End Sub
Public Sub Text1()
If Form7.Text1 = form2.Label3.Caption Then
form2.Label3.Caption = Form7.Text1
End If

End Sub
Public Sub drag()
'If form2.Picture1.DragMode = True Then
'form2.Picture8.Picture = LoadPicture("c:\windows\desktop\vbprojects\war\conn.bmp") '("c:\windows\desktop\vbpojects\war\trash02a.ico ")
If form2.Picture2.DragMode = True Then
form2.Picture8.Picture = LoadPicture("c:\windows\desktop\vbprojects\war\plant.bmp")
End If
'End If

End Sub
