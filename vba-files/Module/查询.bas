Attribute VB_Name = "查询"
Sub SerchByName()
    '
    ' 查询 Macro
    '
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
    Dim filterValue As String
    Dim lastRow1 As Long, lastRow2 As Long, lastRow3 As Long
    Dim i As Long, j As Long, t As Long

    ' 设置工作表
    Set ws1 = ThisWorkbook.Sheets("管理界面") ' 当前工作表
    Set ws2 = ThisWorkbook.Sheets("用户数据") ' 数据源工作表
    Set ws3 = ThisWorkbook.Sheets("资产清单") ' 数据源工作表

    ' 取消工作表保护（绕过密码）
    'ws1.Unprotect Chr(0) ' 使用空字符绕过密码
    ws1.Unprotect Password:="123456" ' 如果未设置密码，可以删除 Password 参数
    ws2.Unprotect Password:="123456" ' 如果未设置密码，可以删除 Password 参数
    ws3.Unprotect Password:="123456" ' 如果未设置密码，可以删除 Password 参数

    ' 获取筛选值（假设下拉菜单在A1单元格）
    filterValue = ws1.Range("F1").Value

    ' 获取Sheet1的最后一行
    lastRow1 = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row + 1

    ' 遍历Sheet2的C列，查找符合条件的数据
    lastRow3 = ws3.Cells(ws3.Rows.Count, 8).End(xlUp).Row

    t = 11

    '删除第2行之后的所有数据
    ws1.Rows("10:" & ws1.Rows.Count).Delete Shift:=xlUp

    ' 复制 Sheet1 的第一行数据
    ws3.Rows(1).Copy

    ' 粘贴到 Sheet2 的第一行
    ws1.Rows(t - 1).PasteSpecial Paste:=xlPasteAll

    For i = 2 To lastRow3 ' 从第2行开始遍历
        If ws3.Cells(i, 3).Value = filterValue Then
            ' 复制符合条件的整行到Sheet1的末尾
            ws3.Rows(i).Copy Destination:=ws1.Rows(t)
            t = t + 1
        End If
    Next i
    If t = 11 Then
        MsgBox "无借用历史！"
    Else
        MsgBox "查询成功！"
    End If
    ' 锁定第二行及以下的所有单元格
    ws1.Rows("1:" & ws1.Rows.Count).Locked = True
    ws2.Rows("1:" & ws2.Rows.Count).Locked = True
    ws3.Rows("1:" & ws3.Rows.Count).Locked = True
    ' 解锁 b1 单元格
    ws1.Range("B1").Locked = False
    ws1.Range("F1").Locked = False
    ws1.Range("B4").Locked = False
    ws1.Range("C4").Locked = False
    ws1.Range("D4").Locked = False
    ws1.Range("E4").Locked = False
    ws1.Range("F4").Locked = False
    ws1.Range("G4").Locked = False
    ws1.Range("H4").Locked = False
    ws1.Range("B7").Locked = False
    ws1.Range("C7").Locked = False
    ws1.Range("D7").Locked = False
    ws1.Range("G7").Locked = False
    ws1.Range("H7").Locked = False
    ws1.Range("I7").Locked = False
    ' 保护工作表，防止修改锁定的单元格
    ws1.Protect Password:="123456", AllowFormattingCells:=True
    ws2.Protect Password:="123456", AllowFormattingCells:=True
    ws3.Protect Password:="123456", AllowFormattingCells:=True
    ThisWorkbook.Save '保存当前工作表
End Sub

Sub SerchImprotant()
    '
    ' 查询入资资产 Macro
    '
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
    Dim filterValue As String
    Dim lastRow1 As Long, lastRow2 As Long, lastRow3 As Long
    Dim i As Long, j As Long, t As Long

    ' 设置工作表
    Set ws1 = ThisWorkbook.Sheets("管理界面") ' 当前工作表
    Set ws2 = ThisWorkbook.Sheets("用户数据") ' 数据源工作表
    Set ws3 = ThisWorkbook.Sheets("资产清单") ' 数据源工作表

    ' 取消工作表保护（绕过密码）
    'ws1.Unprotect Chr(0) ' 使用空字符绕过密码
    ws1.Unprotect Password:="123456" ' 如果未设置密码，可以删除 Password 参数
    ws2.Unprotect Password:="123456" ' 如果未设置密码，可以删除 Password 参数
    ws3.Unprotect Password:="123456" ' 如果未设置密码，可以删除 Password 参数

    ' 获取筛选值（假设下拉菜单在A1单元格）
    filterValue = ws1.Range("F1").Value

    ' 获取Sheet1的最后一行
    lastRow1 = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row + 1

    ' 遍历Sheet2的C列，查找符合条件的数据
    lastRow3 = ws3.Cells(ws3.Rows.Count, 8).End(xlUp).Row

    lastRow2 = ws2.Cells(ws2.Rows.Count, 2).End(xlUp).Row
    t = 11

    '删除第2行之后的所有数据
    ws1.Rows("10:" & ws1.Rows.Count).Delete Shift:=xlUp

    ' 复制 Sheet1 的第一行数据
    ws3.Rows(1).Copy

    ' 粘贴到 Sheet2 的第一行
    ws1.Rows(t - 1).PasteSpecial Paste:=xlPasteAll
    For i = 2 To lastRow3 ' 从第2行开始遍历
        If Not IsEmpty(ws3.Cells(i, 9).Value) Then
            ' 复制符合条件的整行到Sheet1的末尾
            ws3.Rows(i).Copy Destination:=ws1.Rows(t)
            t = t + 1
        End If
    Next i
    If t = 11 Then
        MsgBox "无入资资产！"
    Else
        MsgBox "查询完成！"
    End If
    ' 锁定第二行及以下的所有单元格
    ws1.Rows("1:" & ws1.Rows.Count).Locked = True
    ws2.Rows("1:" & ws2.Rows.Count).Locked = True
    ws3.Rows("1:" & ws3.Rows.Count).Locked = True
    ' 解锁 b1 单元格
    ws1.Range("B1").Locked = False
    ws1.Range("F1").Locked = False
    ws1.Range("B4").Locked = False
    ws1.Range("C4").Locked = False
    ws1.Range("D4").Locked = False
    ws1.Range("E4").Locked = False
    ws1.Range("F4").Locked = False
    ws1.Range("G4").Locked = False
    ws1.Range("H4").Locked = False
    ws1.Range("B7").Locked = False
    ws1.Range("C7").Locked = False
    ws1.Range("D7").Locked = False
    ws1.Range("G7").Locked = False
    ws1.Range("H7").Locked = False
    ws1.Range("I7").Locked = False
    ' 保护工作表，防止修改锁定的单元格
    'ws1.Rows(10 & ":" & t).Locked = False
    ' 保护工作表，防止修改锁定的单元格
    ws1.Protect Password:="123456", AllowFormattingCells:=True
    ws2.Protect Password:="123456", AllowFormattingCells:=True
    ws3.Protect Password:="123456", AllowFormattingCells:=True
    ThisWorkbook.Save '保存当前工作表
End Sub

Sub SerchByType()
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
    Dim filterValue As String
    Dim lastRow1 As Long, lastRow2 As Long, lastRow3 As Long
    Dim i As Long, j As Long, t As Long

    ' 设置工作表
    Set ws1 = ThisWorkbook.Sheets("管理界面") ' 当前工作表
    Set ws2 = ThisWorkbook.Sheets("用户数据") ' 数据源工作表
    Set ws3 = ThisWorkbook.Sheets("资产清单") ' 数据源工作表

    ' 取消工作表保护（绕过密码）
    'ws1.Unprotect Chr(0) ' 使用空字符绕过密码
    ws1.Unprotect Password:="123456" ' 如果未设置密码，可以删除 Password 参数
    ws2.Unprotect Password:="123456" ' 如果未设置密码，可以删除 Password 参数
    ws3.Unprotect Password:="123456" ' 如果未设置密码，可以删除 Password 参数
    ' 获取筛选值（假设下拉菜单在A1单元格）
    filterValue = ws1.Range("B1").Value

    ' 获取Sheet1的最后一行
    lastRow1 = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row + 1

    ' 遍历Sheet2的C列，查找符合条件的数据
    lastRow3 = ws3.Cells(ws3.Rows.Count, 8).End(xlUp).Row

    t = 11

    '删除第2行之后的所有数据
    ws1.Rows("10:" & ws1.Rows.Count).Delete Shift:=xlUp

    ' 复制 Sheet1 的第一行数据
    ws3.Rows(1).Copy

    ' 粘贴到 Sheet2 的第一行
    ws1.Rows(t - 1).PasteSpecial Paste:=xlPasteAll

    For i = 2 To lastRow3 ' 从第2行开始遍历
        If ws3.Cells(i, 8).Value = filterValue Then
            ' 复制符合条件的整行到Sheet1的末尾
            ws3.Rows(i).Copy Destination:=ws1.Rows(t)
            t = t + 1
        End If
    Next i

    ' 锁定第二行及以下的所有单元格
    ws1.Rows("1:" & ws1.Rows.Count).Locked = True
    ws2.Rows("1:" & ws2.Rows.Count).Locked = True
    ws3.Rows("1:" & ws3.Rows.Count).Locked = True
    ' 解锁 b1 单元格
    ws1.Range("B1").Locked = False
    ws1.Range("F1").Locked = False
    ws1.Range("B4").Locked = False
    ws1.Range("C4").Locked = False
    ws1.Range("D4").Locked = False
    ws1.Range("E4").Locked = False
    ws1.Range("F4").Locked = False
    ws1.Range("G4").Locked = False
    ws1.Range("H4").Locked = False
    ws1.Range("B7").Locked = False
    ws1.Range("C7").Locked = False
    ws1.Range("D7").Locked = False
    ws1.Range("G7").Locked = False
    ws1.Range("H7").Locked = False
    ws1.Range("I7").Locked = False
    ' 保护工作表，防止修改锁定的单元格
    ws1.Protect Password:="123456", AllowFormattingCells:=True
    ws2.Protect Password:="123456", AllowFormattingCells:=True
    ws3.Protect Password:="123456", AllowFormattingCells:=True

    ThisWorkbook.Save '保存当前工作表

    MsgBox "查询成功！"
End Sub

