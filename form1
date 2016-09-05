Public sy As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Public Sub rename(a As String, b As Integer)   '读取便签的标题
  Open "C:\" & a For Input As #2
  Line Input #2, DEMO
  SSTab1.TabCaption(b - 1) = SSTab1.TabCaption(b - 1) & DEMO & vbCrLf
  Close #2
End Sub

Public Sub content(a As String, b As Integer)    '读取便签的内容
  Open "C:\" & a For Input As #1
  While Not EOF(1)
    Line Input #1, DEMO
    Text1(b).Text = Text1(b).Text & DEMO & vbCrLf
  Wend
  Close #1
End Sub

Public Sub save(a As String, b As Integer, c As String)      '保存便签的内容及标题
  Open "C:\" & a For Output As #b
  Print #b, c
  Close #b
End Sub

Public Sub color(s As String)                       '修改颜色函数
  Dim k As Long
  For k = 1 To 3                          '3
    Text1(k).ForeColor = s
  Next k
End Sub

Private Sub Combo1_Click()
  Text1(1).FontSize = Combo1.Text
  Text1(2).FontSize = Combo1.Text
  Text1(3).FontSize = Combo1.Text
End Sub

Private Sub Combo2_Click()
  Select Case Combo2.Text
    Case "黑色": Call color(vbBlack)
    Case "红色": Call color(vbRed)
    Case "蓝色": Call color(vbBlue)
    Case "青色": Call color(vbCyan)
    Case "洋红": Call color(vbMagenta)
    Case "黄色": Call color(vbYellow)
End Select
End Sub

Private Sub Command1_Click()
  Dim i As Integer
  For i = 1 To 3
    Text1(i).FontBold = Not Text1(i).FontBold
  Next i
End Sub

Private Sub Command2_Click()
  Dim i As Integer
  For i = 1 To 3
    Text1(i).FontItalic = Not Text1(i).FontItalic
  Next i
End Sub

Private Sub Command3_Click()
  ShellExecute hWnd, "open", "http://shop111837330.taobao.com/", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub Command4_Click()
  Form1.Height = 7680
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim j As Integer
  For j = 1 To 3      '1
    Call save(j & ".txt", j, Text1(j).Text)
    Call save("caption" & j & ".txt", j + 1, SSTab1.TabCaption(j - 1))
  Next j

  Open "C:\老板便签数据文件.txt" For Output As #3         '保存文本字号  此处可考虑优化
  Print #3, Text1(1).FontSize
  Close #3

  Open "C:\bq2.txt" For Output As #4                         '保存文本颜色   此处可考虑优化
  Print #4, Text1(1).ForeColor
  Close #4
End Sub

Private Sub Form_Load()
  Data = "C:\r8.txt"
  If Dir(Data) = "" Then
    Open "C:\r8.txt" For Output As #3
    Close #3
    Text1(1).Text = "说明：" & vbCrLf & "1.双击标签可以重命名" & vbCrLf & vbCrLf & "2.便签内容是自动保存的" & vbCrLf & vbCrLf & "3.感谢您的下载使用，有任何问题或建议欢迎Q我。" & vbCrLf & "360有时会误报，我也是醉了。。" & vbCrLf & vbCrLf & "4.更多功能下次更新"
  Else
    Dim i As Integer
    For i = 1 To 3            '2
      Call content(i & ".txt", i)
      SSTab1.TabCaption(i - 1) = ""
      Call rename("caption" & i & ".txt", i)
    Next i

    Open "C:\老板便签数据文件.txt" For Input As #3        '读取文本字号
    Line Input #3, DEMO
    For i = 1 To 3
      Text1(i).FontSize = DEMO & vbCrLf
    Next i
    Close #3

    Open "C:\bq2.txt" For Input As #4                       '读取文本颜色
    Line Input #4, DEMO2
    For i = 1 To 3
      Text1(i).ForeColor = DEMO2 & vbCrLf
    Next i
    Close #4
  End If
End Sub

Private Sub SSTab1_dblClick()  '双击修改便签标题
  If SSTab1.Tab = 0 Then
    SSTab1.TabCaption(0) = InputBox("请输入便签名~", "提示", SSTab1.TabCaption(0))
  ElseIf SSTab1.Tab = 1 Then
    SSTab1.TabCaption(1) = InputBox("请输入便签名~", "提示", SSTab1.TabCaption(1))
  ElseIf SSTab1.Tab = 2 Then
    SSTab1.TabCaption(2) = InputBox("请输入便签名~", "提示", SSTab1.TabCaption(2))
  End If
End Sub

'若要增加便签数量，需增加一个text1()的元素，修改tabcaption的tab属性，在相关目录创建好一个便签内容文档，一个便签名文档，然后修改1和2和3处的循环次数即可。
'V1.0 完成时间2015年3月24日21:54:32，下次可改进方向：1.按钮增加、删除便签数量；  2.按钮更改界面风格（背景切换、TEXT字体大小、颜色等）

