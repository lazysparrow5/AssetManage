Attribute VB_Name = "���"
Sub AddEquipment()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim maxRow As Long
    Dim maxValue As Integer
    Dim i As Long
    Dim i2 As Long
    Dim t1 As Long

    ' ���ù����������ǻ������
    Set ws = ThisWorkbook.Sheets("�ʲ��嵥") ' ��ǰ������

    ' �ҵ����һ��
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
    Set ws1 = ThisWorkbook.Sheets("�������") ' ��ǰ������

    Dim cell4 As Range
    Set cell4 = ws1.Range("B4") ' Ҫ���ĵ�Ԫ��

    Dim ws2 As Worksheet
    Set ws2 = ThisWorkbook.Sheets("�ʲ��嵥") ' ��ǰ������
    Dim ws3 As Worksheet
    Set ws3 = ThisWorkbook.Sheets("�û�����") ' ��ǰ������
    ws1.Unprotect Password:="123456" ' ���δ�������룬����ɾ�� Password ����
    ws2.Unprotect Password:="123456" ' ���δ�������룬����ɾ�� Password ����
    ws3.Unprotect Password:="123456" ' ���δ�������룬����ɾ�� Password ����
    If IsEmpty(cell) Or cell.Value = "" Then
        MsgBox "��Ʒ���Ʋ���Ϊ�գ�", vbExclamation
    ElseIf Not IsTextDashNumberFormat(cell.Value) Then
        MsgBox " ��Ʒ���Ʋ�����'�ı�-����'��ʽ", vbInformation
    Else

    End If

    t1 = 1
    For i2 = 2 To lastRow
        If cell.Value = ws.Cells(i2, 2).Value Then
            MsgBox " ��Ʒ�����Ѵ���", vbInformation
            t1 = 10
         Exit For  ' �ҵ�Ŀ��������˳�ѭ��
        End If
    Next i2
    If Not t1 = 10 Then
        If IsEmpty(cell1) Or cell1.Value = "" Then
            MsgBox "��Ʒ���Ͳ���Ϊ�գ�", vbExclamation
            '        MsgBox "�����֤ͨ����" & cell.Value, vbInformation
            'Elseif IsEmpty(cell2) Or cell2.Value = "" Then
            ' MsgBox "�����˲���Ϊ�գ�", vbExclamation
        ElseIf IsEmpty(cell3) Or cell3.Value = "" Then
            MsgBox "�洢λ�ò���Ϊ�գ�", vbExclamation
        Else
            ' ����Ƿ����㹻��������
            If lastRow < 2 Then Exit Sub ' ֻ�б�ͷ��ձ�

                ' ��ʼ�����ֵ
                maxValue = Val(ws.Cells(2, 1).Value)
                maxRow = 2

                ' �ӵڶ��п�ʼ����������
                For i = 3 To lastRow
                    If Val(ws.Cells(i, 1).Value) > maxValue Then
                        maxValue = Val(ws.Cells(i, 1).Value)
                        maxRow = i
                    End If
                Next i

                ' ��������·���������
                ws.Rows(maxRow + 1).Insert Shift:=xlDown

                ' ���Ƹ�ʽ������
                ws.Rows(maxRow).Copy
                ws.Rows(maxRow + 1).PasteSpecial Paste:=xlPasteFormats
                ws.Rows(maxRow + 1).PasteSpecial Paste:=xlPasteValues

                ' �������
                ws.Cells(maxRow + 1, 1).Value = maxValue + 1

                ' ��ȡ B7 �� C7 ��ֵ
                searchValue = ws1.Range("B4").Value
                fillValue = ws1.Range("C4").Value
                description1 = ws1.Range("D4").Value
                code1 = ws1.Range("E4").Value
                code2 = ws1.Range("F4").Value
                code3 = ws1.Range("G4").Value
                code4 = ws1.Range("H4").Value

                ' �ҵ�ƥ���У���� C ��
                ws.Cells(maxRow + 1, 2).Value = searchValue
                ws.Cells(maxRow + 1, 8).Value = fillValue
                ' ��䵱ǰʱ�䵽 E �У���ʽ��Ϊ "YYYY-MM-DD HH:MM:SS"��
                ws.Cells(maxRow + 1, 13).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
                ws.Cells(maxRow + 1, 3).Value = description1
                ws.Cells(maxRow + 1, 9).Value = code1
                ws.Cells(maxRow + 1, 10).Value = code2
                ws.Cells(maxRow + 1, 5).Value = code3
                ws.Cells(maxRow + 1, 4).Value = code4
                '�û���
                computerName = Environ("COMPUTERNAME")
                If computerName = "PC1026871" Then
                    ws.Cells(maxRow + 1, 6).Value = "���ͻ�"
                ElseIf computerName = "PC1028363" Then
                    ws.Cells(maxRow + 1, 6).Value = "�´�"
                ElseIf computerName = "PC1027303" Then
                    ws.Cells(maxRow + 1, 6).Value = "����"
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
                MsgBox "���ɹ������ʲ��嵥�鿴", vbExclamation
            End If
        End If
        ' ���������
        Application.CutCopyMode = False
        ' �����ڶ��м����µ����е�Ԫ��
        ws1.Rows("1:" & ws1.Rows.Count).Locked = True
        ws2.Rows("1:" & ws2.Rows.Count).Locked = True
        ws3.Rows("1:" & ws3.Rows.Count).Locked = True
        ' ���� b1 ��Ԫ��
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
        ' ������������ֹ�޸������ĵ�Ԫ��
        ws1.Protect Password:="123456", AllowFormattingCells:=True
        ws2.Protect Password:="123456", AllowFormattingCells:=True
        ws3.Protect Password:="123456", AllowFormattingCells:=True
        ThisWorkbook.Save '���浱ǰ������
End Sub

Function IsTextDashNumberFormat(cellValue As String) As Boolean
    Dim parts() As String

    ' ��"-"����ַ���
    parts = Split(cellValue, "-")

    ' �������÷ֳ������֣��ҵڶ�����������
    If UBound(parts) = 1 Then
        If IsNumeric(parts(1)) Then
            IsTextDashNumberFormat = True
         Exit Function
        End If
    End If

    IsTextDashNumberFormat = False
End Function

