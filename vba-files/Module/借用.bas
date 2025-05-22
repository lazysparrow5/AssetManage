Attribute VB_Name = "����"

Private IndexBorrowValue As String
Private UserBorrowValue As String
Private BriefBorrowValue As String

Sub UserBorrow()
    '
    ' �������
    '
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
    ' ���ù�����Ĭ��ʹ�õ�ǰ�������
    Set ws = ActiveSheet
    Dim ws1 As Worksheet
    ' ���ù�����
    Set ws1 = ThisWorkbook.Sheets("�������") ' ��ǰ������

    '�ʲ��嵥�����
    Dim ws2 As Worksheet
    Dim lastRow2 As Long
    Dim j As Long
    Dim matchFound1 As Boolean
    Set ws2 = ThisWorkbook.Sheets("�ʲ��嵥") ' ��ǰ������
    Dim ws3 As Worksheet
    Set ws3 = ThisWorkbook.Sheets("�û�����") ' ��ǰ������

'******************************************************************************
    IndexBorrowValue = ManageSheet.Range(INDEX_BORROW_CELL).Value
    UserBorrowValue = ManageSheet.Range(USER_BORROW_CELL).Value
    BriefBorrowValue = ManageSheet.Range(BRIEF_BORROW_CELL).Value

    If Not BorrowInputValid Then
        Exit Sub  
    End If

    DataUpdate



'******************************************************************************


    Dim flag As Long
    flag = 1

        '�ʲ��嵥����
        ' �����ʲ��嵥 A �����һ��
        lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
        matchFound1 = False
        For j = 1 To lastRow2
            If ws2.Cells(j, 1).Value = searchValue Then

                If Not IsEmpty(ws2.Cells(j, 3).Value) Then
                    MsgBox "��Ʒ�ѱ�����,���ȹ黹��", vbExclamation
                Else
                    flag = 10
                    ' �ҵ�ƥ���У���� C ��
                    ws2.Cells(j, 3).Value = fillValue
                    ws2.Cells(j, 4).Value = ""
                    'ws2.Cells(j, 5).Value = description1
                    ' ��䵱ǰʱ�䵽 E �У���ʽ��Ϊ "YYYY-MM-DD HH:MM:SS"��
                    ws2.Cells(j, 13).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")

                    If description1 <> "" Then
                        ws2.Cells(j, 5).Value = description1
                    End If

                    '�û���
                    computerName = Environ("COMPUTERNAME")
                    If computerName = "PC1026871" Then
                        ws2.Cells(j, 6).Value = "���ͻ�"
                    ElseIf computerName = "PC1028363" Then
                        ws2.Cells(j, 6).Value = "�´�"
                    ElseIf computerName = "PC1027303" Then
                        ws2.Cells(j, 6).Value = "����"
                    Else
                        ws2.Cells(j, 6).Value = computerName
                    End If
                    matchFound1 = True
                End If
            End If
        Next j

        If flag = 10 Then

            ' ���� A �����һ��
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            ' ���� A �в���ƥ����
            matchFound = False
            For i = 1 To lastRow
                If ws.Cells(i, 1).Value = searchValue Then
                    ' �ҵ�ƥ���У���� C ��
                    ws.Cells(i, 3).Value = fillValue
                    'ws.Cells(i, 4).Value = ""

                    If description1 <> "" Then
                        ws.Cells(i, 5).Value = description1
                    End If

                    ' ��䵱ǰʱ�䵽 E �У���ʽ��Ϊ "YYYY-MM-DD HH:MM:SS"��
                    ws.Cells(i, 13).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
                    '�û���
                    computerName = Environ("COMPUTERNAME")
                    If computerName = "PC1026871" Then
                        ws.Cells(i, 6).Value = "���ͻ�"
                    ElseIf computerName = "PC1028363" Then
                        ws.Cells(i, 6).Value = "�´�"
                    ElseIf computerName = "PC1027303" Then
                        ws.Cells(i, 6).Value = "����"
                    Else
                        ws.Cells(i, 6).Value = computerName
                    End If
                    matchFound = True
                End If
            Next i

            ws.Range("B7").Value = ""
            ws.Range("C7").Value = ""
            ws.Range("D7").Value = ""
            MsgBox "���óɹ�", vbExclamation
        End If

End Sub

Sub UserReturn()
    '
    ' ����黹 Macro
    '
    Dim computerName As String
    '  ���Գ���
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
    ' ���ù�����Ĭ��ʹ�õ�ǰ�������
    Set ws = ActiveSheet
    Dim ws1 As Worksheet
    ' ���ù�����
    Set ws1 = ThisWorkbook.Sheets("�������") ' ��ǰ������


    '�ʲ��嵥�����
    Dim ws2 As Worksheet
    Dim lastRow2 As Long
    Dim j As Long
    Dim matchFound1 As Boolean
    Set ws2 = ThisWorkbook.Sheets("�ʲ��嵥") ' ��ǰ������

    Dim ws3 As Worksheet
    Set ws3 = ThisWorkbook.Sheets("�û�����") ' ��ǰ������

    ws1.Unprotect Password:="123456" ' ���δ�������룬����ɾ�� Password ����
    ws2.Unprotect Password:="123456" ' ���δ�������룬����ɾ�� Password ����
    ws3.Unprotect Password:="123456" ' ���δ�������룬����ɾ�� Password ����
    computerName = Environ("COMPUTERNAME")

    If computerName <> "PC1026871" And computerName <> "PC1028363" Then
        MsgBox "���ҹ���Ա-�´Ϲ黹��", vbExclamation
    ElseIf IsEmpty(cell) Or cell.Value = "" Then
        MsgBox "��Ų���Ϊ�գ�", vbExclamation
    ElseIf Not IsNumeric(cell.Value) Then
        MsgBox "��ű���Ϊ���֣�", vbCritical
    ElseIf IsEmpty(cell1) Or cell.Value = "" Then
        MsgBox "��ַ����Ϊ�գ�", vbExclamation
        '        MsgBox "�����֤ͨ����" & cell.Value, vbInformation

    Else
        ' ��ȡ B7 �� C7 ��ֵ
        searchValue = ws.Range("G7").Value
        fillValue = ws.Range("H7").Value
        description1 = ws.Range("I7").Value
        ' ���ҹ������ A �����һ��
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        ' �����ʲ��嵥 A �����һ��
        lastRow2 = ws2.Cells(ws2.Rows.Count, "A").End(xlUp).Row
        ' ���� A �в���ƥ����
        matchFound = False
        For i = 1 To lastRow
            If ws.Cells(i, 1).Value = searchValue Then
                ' �ҵ�ƥ���У���� C ��
                ws.Cells(i, 3).Value = ""
                ws.Cells(i, 4).Value = fillValue
                ' ��䵱ǰʱ�䵽 E �У���ʽ��Ϊ "YYYY-MM-DD HH:MM:SS"��
                ws.Cells(i, 13).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")

                If description1 <> "" Then
                    ws.Cells(i, 5).Value = description1
                End If
                '�û���
                If computerName = "PC1026871" Then
                    ws.Cells(i, 6).Value = "���ͻ�"
                ElseIf computerName = "PC1028363" Then
                    ws.Cells(i, 6).Value = "�´�"
                ElseIf computerName = "PC1027303" Then
                    ws.Cells(i, 6).Value = "����"
                Else
                    ws.Cells(i, 6).Value = computerName
                End If
                matchFound = True
            End If
        Next i
        '�ʲ��嵥����
        matchFound1 = False
        For j = 1 To lastRow2
            If ws2.Cells(j, 1).Value = searchValue Then
                ' �ҵ�ƥ���У���� C ��
                ws2.Cells(j, 3).Value = ""
                ws2.Cells(j, 4).Value = fillValue
                ' ��䵱ǰʱ�䵽 E �У���ʽ��Ϊ "YYYY-MM-DD HH:MM:SS"��
                ws2.Cells(j, 13).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")
                If description1 <> "" Then
                    ws2.Cells(j, 5).Value = description1
                End If
                '�û���
                computerName = Environ("COMPUTERNAME")
                If computerName = "PC1026871" Then
                    ws2.Cells(j, 6).Value = "���ͻ�"
                ElseIf computerName = "PC1028363" Then
                    ws2.Cells(j, 6).Value = "�´�"
                ElseIf computerName = "PC1027303" Then
                    ws2.Cells(j, 6).Value = "����"
                Else
                    ws2.Cells(j, 6).Value = computerName
                End If
                matchFound1 = True
            End If
        Next j
        ws.Range("G7").Value = ""
        ws.Range("H7").Value = ""
        ws.Range("I7").Value = ""
        MsgBox "�黹�ɹ�", vbExclamation
    End If
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

Private Function BorrowInputValid() As Boolean
    Dim index As Variant
    BorrowInputValid = False

    index = SafeStringToLong(IndexValue)
    Select Case True
    Case Trim(IndexValue) = ""
        MsgBox "��Ų���Ϊ��"
    Case Not IsNumeric(IndexValue)
        MsgBox "��ű���Ϊ����"
    Case IndexValue = Fix(IndexValue)
        MsgBox "��ű���Ϊ����"
    Case index > 0 & index < AssetsIndexMax
        MsgBox "��Ų��ڷ�Χ��"
    Case Trim(UserBorrowValue) = ""
        MsgBox "�����˲���Ϊ��"
    Case Else
        BorrowInputValid = True
    End Select

End Function


