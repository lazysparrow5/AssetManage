Attribute VB_Name = "添加"
Sub AddEquipment()

    If UserName <> AdminID Then
        MsgBox "归还处于维护中，请联系管理员归还"
        Exit Sub
    End If

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim maxRow As Long
    Dim maxValue As Integer
    Dim i As Long
    Dim i2 As Long
    Dim t1 As Long

    ' 设置工作表（假设是活动工作表）
    Set ws = ThisWorkbook.Sheets("资产清单") ' 当前工作表

    ' 找到最后一行
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim computerName As String
    Dim cell As Range
    Dim cell1 As Range
    Dim cell2 As Range
    Dim cell3 As Range
    Set cell = Range("B4")
    Set cell1 = Range("C4")
    Set cell2 = Range("D4")
    Set cell3 = Range("H4")

    Dim searchValue As String
    Dim fillValue As String
    Dim description1 As String
    Dim code1 As String
    Dim code2 As String
    Dim code3 As String
    Dim code4 As String

    Dim ws1 As Worksheet
    Set ws1 = ThisWorkbook.Sheets("管理界面") ' 当前工作表

    Dim cell4 As Range
    Set cell4 = ws1.Range("B4") ' 要检查的单元格

    Dim ws2 As Worksheet
    Set ws2 = ThisWorkbook.Sheets("资产清单") ' 当前工作表
    Dim ws3 As Worksheet
    Set ws3 = ThisWorkbook.Sheets("用户数据") ' 当前工作表
    ws1.Unprotect Password:="123456" ' 如果未设置密码，可以删除 Password 参数
    ws2.Unprotect Password:="123456" ' 如果未设置密码，可以删除 Password 参数
    ws3.Unprotect Password:="123456" ' 如果未设置密码，可以删除 Password 参数
    If IsEmpty(cell) Or cell.Value = "" Then
        MsgBox "物品名称不能为空！", vbExclamation
    ElseIf Not IsTextDashNumberFormat(cell.Value) Then
        MsgBox " 物品名称不符合'文本-数字'格式", vbInformation
    Else

    End If

    t1 = 1
    For i2 = 2 To lastRow
        If cell.Value = ws.Cells(i2, 2).Value Then
            MsgBox " 物品名称已存在", vbInformation
            t1 = 10
         Exit For  ' 找到目标后立即退出循环
        End If
    Next i2
    If Not t1 = 10 Then
        If IsEmpty(cell1) Or cell1.Value = "" Then
            MsgBox "物品类型不能为空！", vbExclamation
            '        MsgBox "序号验证通过：" & cell.Value, vbInformation
            'Elseif IsEmpty(cell2) Or cell2.Value = "" Then
            ' MsgBox "借用人不能为空！", vbExclamation
        ElseIf IsEmpty(cell3) Or cell3.Value = "" Then
            MsgBox "存储位置不能为空！", vbExclamation
        Else
            ' 检查是否有足够的数据行
            If lastRow < 2 Then Exit Sub ' 只有表头或空表

                ' 初始化最大值
                maxValue = Val(ws.Cells(2, 1).Value)
                maxRow = 2

                ' 从第二行开始查找最大序号
                For i = 3 To lastRow
                    If Val(ws.Cells(i, 1).Value) > maxValue Then
                        maxValue = Val(ws.Cells(i, 1).Value)
                        maxRow = i
                    End If
                Next i

                ' 在最大行下方插入新行
                ws.Rows(maxRow + 1).Insert Shift:=xlDown

                ' 复制格式和数据
                ws.Rows(maxRow).Copy
                ws.Rows(maxRow + 1).PasteSpecial Paste:=xlPasteFormats
                ws.Rows(maxRow + 1).PasteSpecial Paste:=xlPasteValues

                ' 更新序号
                ws.Cells(maxRow + 1, 1).Value = maxValue + 1

                ' 获取 B7 和 C7 的值
                searchValue = ws1.Range("B4").Value
                fillValue = ws1.Range("C4").Value
                description1 = ws1.Range("D4").Value
                code1 = ws1.Range("E4").Value
                code2 = ws1.Range("F4").Value
                code3 = ws1.Range("G4").Value
                code4 = ws1.Range("H4").Value

                ' 找到匹配行，填充 C 列
                ws.Cells(maxRow + 1, 2).Value = searchValue
                ws.Cells(maxRow + 1, 8).Value = fillValue
                ' 填充当前时间到 E 列（格式化为 "YYYY-MM-DD HH:MM:SS"）
                ws.Cells(maxRow + 1, 13).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
                ws.Cells(maxRow + 1, 3).Value = description1
                ws.Cells(maxRow + 1, 9).Value = code1
                ws.Cells(maxRow + 1, 10).Value = code2
                ws.Cells(maxRow + 1, 5).Value = code3
                ws.Cells(maxRow + 1, 4).Value = code4
                '用户名
                computerName = Environ("COMPUTERNAME")
                If computerName = "PC1026871" Then
                    ws.Cells(maxRow + 1, 6).Value = "唐焱辉"
                ElseIf computerName = "PC1028363" Then
                    ws.Cells(maxRow + 1, 6).Value = "陈聪"
                ElseIf computerName = "PC1027303" Then
                    ws.Cells(maxRow + 1, 6).Value = "廖雷"
                Else
                    ws.Cells(maxRow + 1, 6).Value = computerName
                End If
                ws.Cells(maxRow + 1, 7).Value = computerName
                ws1.Range("B4").Value = ""
                ws1.Range("C4").Value = ""
                ws1.Range("D4").Value = ""
                ws1.Range("E4").Value = ""
                ws1.Range("F4").Value = ""
                ws1.Range("G4").Value = ""
                ws1.Range("H4").Value = ""
                MsgBox "入库成功，到资产清单查看", vbExclamation
            End If
        End If
        ' 清除剪贴板
        Application.CutCopyMode = False
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

Function IsTextDashNumberFormat(cellValue As String) As Boolean
    Dim parts() As String

    ' 按"-"拆分字符串
    parts = Split(cellValue, "-")

    ' 必须正好分成两部分，且第二部分是数字
    If UBound(parts) = 1 Then
        If IsNumeric(parts(1)) Then
            IsTextDashNumberFormat = True
         Exit Function
        End If
    End If

    IsTextDashNumberFormat = False
End Function

