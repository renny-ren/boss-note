Private Sub Command1_Click()
  MsgBox ("请把“TABCTL32.OCX”这个文件放到“C:\Windows\SysWOW64”目录下即可！ " & vbCrLf & "如果是32位的系统就放到C:\Windows\System32目录下" & vbCrLf & "另外，360可能会误报O__O …，请放心使用")
End Sub

Private Sub Form_Load()
  Data = "C:\r8.txt"
    If Dir(Data) = "" Then
    Else
      Form1.Show
      Unload Me
    End If
End Sub

Private Sub Image2_Click()
  Form1.Show
  Unload Me
End Sub
