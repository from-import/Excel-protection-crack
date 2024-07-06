# Excel-protection-crack
破解受保护的Excel



## 方法

    打开Excel文件：确保该文件已启用宏。
    打开VBA编辑器：
        按 Alt + F11 打开VBA编辑器。
    插入新模块：
        在VBA编辑器中，右键点击任意工作簿，然后选择插入 -> 模块。
    复制并粘贴代码：

        将以下代码粘贴到新模块中：

        vba

    Sub UnprotectSheet()
        Dim i As Integer, j As Integer, k As Integer
        Dim l As Integer, m As Integer, n As Integer
        Dim i1 As Integer, i2 As Integer, i3 As Integer
        Dim i4 As Integer, i5 As Integer, i6 As Integer
        On Error Resume Next
        For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
        For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
        For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
        For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126
        ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & _
        Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
        Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
        If ActiveSheet.ProtectContents = False Then
        MsgBox "Password is " & Chr(i) & Chr(j) & Chr(k) & _
        Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
        Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
        Exit Sub
        End If
        Next: Next: Next: Next: Next: Next
        Next: Next: Next: Next: Next: Next
    End Sub

    按 F5 或点击“运行”按钮执行代码。
