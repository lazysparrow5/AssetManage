Attribute VB_Name = "借用"
Sub UserBorrow()
    '
    ' 点击借用 Macro
    '
    Dim computerName As String

    Dim cell As Range
    Dim cell1 As Range
    Set cell = Range("B7")
    Set cell1 = Range("C7")

    Dim ws As Worksheet
    Dim searchValue As String
    Dim fillValue As String
    Dim description1 As String
    Dim lastRow As Long
    Dim i As Long
    Dim matchFound As Boolean
    ' 设置工作表（默认使用当前活动工作表）
    Set ws = ActiveSheet
    Dim ws1 As Worksheet
    ' 设置工作表
    Set ws1 = ThisWorkbook.Sheets("管理界面") ' 当前工作表

    '资产清单表变量
    Dim ws2 As Worksheet
    Dim lastRow2 As Long
    Dim j As Long
    Dim matchFound1 As Boolean
    Set ws2 = ThisWorkbook.Sheets("资产清单") ' 当前工作表
    Dim ws3 As Worksheet
    Set ws3 = ThisWorkbook.Sheets("用户数据") ' 当前工作表

    ws1.Unprotect Password:="123456" ' 如果未设置密码，可以删除 Password 参数
    ws2.Unprotect Password:="123456" ' 如果未设置密码，可以删除 Password 参数
    ws3.Unprotect Password:="123456" ' 如果未设置密码，可以删除 Password 参数


    Dim flag As Long
    flag = 1

    If IsEmpty(cell) Or cell.Value = "" Then
        MsgBox "序号不能为空！", vbExclamation
    ElseIf Not IsNumeric(cell.Value) Then
        MsgBox "序号必须为数字！", vbCritical
    ElseIf IsEmpty(cell1) Or cell.Value = "" Then
        MsgBox "借用人不能为空！", vbExclamation
        '        MsgBox "序号验证通过：" & cell.Value, vbInformation
    Else
        ' 获取 B7 和 C7 的值
        searchValue = ws.Range("B7").Value
        fillValue = ws.Range("C7").Value
        description1 = ws.Range("D7").Value



        '资产清单表处理
        ' 查找资产清单 A 列最后一行
        lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
        matchFound1 = False
        For j = 1 To lastRow2
            If ws2.Cells(j, 1).Value = searchValue Then

                If Not IsEmpty(ws2.Cells(j, 3).Value) Then
                    MsgBox "物品已被借用,请先归还！", vbExclamation
                Else
                    flag = 10
                    ' 找到匹配行，填充 C 列
                    ws2.Cells(j, 3).Value = fillValue
                    ws2.Cells(j, 4).Value = ""
                    'ws2.Cells(j, 5).Value = description1
                    ' 填充当前时间到 E 列（格式化为 "YYYY-MM-DD HH:MM:SS"）
                    ws2.Cells(j, 13).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")

                    If description1 <> "" Then
                        ws2.Cells(j, 5).Value = description1
                    End If

                    '用户名
                    computerName = Environ("COMPUTERNAME")
                    If computerName = "PC1026871" Then
                        ws2.Cells(j, 6).Value = "唐焱辉"
                    ElseIf computerName = "PC1028363" Then
                        ws2.Cells(j, 6).Value = "陈聪"
                    ElseIf computerName = "PC1027303" Then
                        ws2.Cells(j, 6).Value = "廖雷"
                    Else
                        ws2.Cells(j, 6).Value = computerName
                    End If
                    matchFound1 = True
                End If
            End If
        Next j

        If flag = 10 Then

            ' 查找 A 列最后一行
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            ' 遍历 A 列查找匹配项
            matchFound = False
            For i = 1 To lastRow
                If ws.Cells(i, 1).Value = searchValue Then
                    ' 找到匹配行，填充 C 列
                    ws.Cells(i, 3).Value = fillValue
                    'ws.Cells(i, 4).Value = ""

                    If description1 <> "" Then
                        ws.Cells(i, 5).Value = description1
                    End If

                    ' 填充当前时间到 E 列（格式化为 "YYYY-MM-DD HH:MM:SS"）
                    ws.Cells(i, 13).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
                    '用户名
                    computerName = Environ("COMPUTERNAME")
                    If computerName = "PC1026871" Then
                        ws.Cells(i, 6).Value = "唐焱辉"
                    ElseIf computerName = "PC1028363" Then
                        ws.Cells(i, 6).Value = "陈聪"
                    ElseIf computerName = "PC1027303" Then
                        ws.Cells(i, 6).Value = "廖雷"
                    Else
                        ws.Cells(i, 6).Value = computerName
                    End If
                    matchFound = True
                End If
            Next i

            ws.Range("B7").Value = ""
            ws.Range("C7").Value = ""
            ws.Range("D7").Value = ""
            MsgBox "借用成功", vbExclamation
        End If
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

Sub UserReturn()
    '
    ' 点击归还 Macro
    '
    Dim computerName As String
    '  测试程序
    Dim cell As Range
    Dim cell1 As Range
    Set cell = Range("G7")
    Set cell1 = Range("H7")


    Dim ws As Worksheet
    Dim searchValue As String
    Dim fillValue As String
    Dim description1 As String
    Dim lastRow As Long
    Dim i As Long
    Dim matchFound As Boolean
    ' 设置工作表（默认使用当前活动工作表）
    Set ws = ActiveSheet
    Dim ws1 As Worksheet
    ' 设置工作表
    Set ws1 = ThisWorkbook.Sheets("管理界面") ' 当前工作表


    '资产清单表变量
    Dim ws2 As Worksheet
    Dim lastRow2 As Long
    Dim j As Long
    Dim matchFound1 As Boolean
    Set ws2 = ThisWorkbook.Sheets("资产清单") ' 当前工作表

    Dim ws3 As Worksheet
    Set ws3 = ThisWorkbook.Sheets("用户数据") ' 当前工作表

    ws1.Unprotect Password:="123456" ' 如果未设置密码，可以删除 Password 参数
    ws2.Unprotect Password:="123456" ' 如果未设置密码，可以删除 Password 参数
    ws3.Unprotect Password:="123456" ' 如果未设置密码，可以删除 Password 参数
    computerName = Environ("COMPUTERNAME")

    If computerName <> "PC1026871" And computerName <> "PC1028363" Then
        MsgBox "请找管理员-陈聪归还！", vbExclamation
    ElseIf IsEmpty(cell) Or cell.Value = "" Then
        MsgBox "序号不能为空！", vbExclamation
    ElseIf Not IsNumeric(cell.Value) Then
        MsgBox "序号必须为数字！", vbCritical
    ElseIf IsEmpty(cell1) Or cell.Value = "" Then
        MsgBox "地址不能为空！", vbExclamation
        '        MsgBox "序号验证通过：" & cell.Value, vbInformation

    Else
        ' 获取 B7 和 C7 的值
        searchValue = ws.Range("G7").Value
        fillValue = ws.Range("H7").Value
        description1 = ws.Range("I7").Value
        ' 查找管理界面 A 列最后一行
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        ' 查找资产清单 A 列最后一行
        lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
        ' 遍历 A 列查找匹配项
        matchFound = False
        For i = 1 To lastRow
            If ws.Cells(i, 1).Value = searchValue Then
                ' 找到匹配行，填充 C 列
                ws.Cells(i, 3).Value = ""
                ws.Cells(i, 4).Value = fillValue
                ' 填充当前时间到 E 列（格式化为 "YYYY-MM-DD HH:MM:SS"）
                ws.Cells(i, 13).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")

                If description1 <> "" Then
                    ws.Cells(i, 5).Value = description1
                End If
                '用户名
                If computerName = "PC1026871" Then
                    ws.Cells(i, 6).Value = "唐焱辉"
                ElseIf computerName = "PC1028363" Then
                    ws.Cells(i, 6).Value = "陈聪"
                ElseIf computerName = "PC1027303" Then
                    ws.Cells(i, 6).Value = "廖雷"
                Else
                    ws.Cells(i, 6).Value = computerName
                End If
                matchFound = True
            End If
        Next i
        '资产清单表处理
        matchFound1 = False
        For j = 1 To lastRow2
            If ws2.Cells(j, 1).Value = searchValue Then
                ' 找到匹配行，填充 C 列
                ws2.Cells(j, 3).Value = ""
                ws2.Cells(j, 4).Value = fillValue
                ' 填充当前时间到 E 列（格式化为 "YYYY-MM-DD HH:MM:SS"）
                ws2.Cells(j, 13).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
                If description1 <> "" Then
                    ws2.Cells(j, 5).Value = description1
                End If
                '用户名
                computerName = Environ("COMPUTERNAME")
                If computerName = "PC1026871" Then
                    ws2.Cells(j, 6).Value = "唐焱辉"
                ElseIf computerName = "PC1028363" Then
                    ws2.Cells(j, 6).Value = "陈聪"
                ElseIf computerName = "PC1027303" Then
                    ws2.Cells(j, 6).Value = "廖雷"
                Else
                    ws2.Cells(j, 6).Value = computerName
                End If
                matchFound1 = True
            End If
        Next j
        ws.Range("G7").Value = ""
        ws.Range("H7").Value = ""
        ws.Range("I7").Value = ""
        MsgBox "归还成功", vbExclamation
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


