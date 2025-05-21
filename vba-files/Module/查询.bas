Attribute VB_Name = "��ѯ"

Const TYPE_FILTER_CELL As String = "B1"        ' ��������λ��
Const NAME_FILTER_CELL As String = "F1"     ' ��������λ��
Const TARGET_ROW_START As Long = 11       ' Ŀ����ʼ��
Const TYPE_COLUMN As Long = 8             ' �豸�����к�

Sub SerchByName()
    '
    ' ��ѯ Macro
    '
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
    Dim filterValue As String
    Dim lastRow1 As Long, lastRow2 As Long, lastRow3 As Long
    Dim i As Long, j As Long, t As Long

    ' ���ù�����
    Set ws1 = ThisWorkbook.Sheets("�������") ' ��ǰ������
    Set ws2 = ThisWorkbook.Sheets("�������") ' ��ǰ������
    Set ws2 = ThisWorkbook.Sheets("�û�����") ' ����Դ������
    Set ws3 = ThisWorkbook.Sheets("�ʲ��嵥") ' ����Դ������

    ' ȡ�������������ƹ����룩
    'ws1.Unprotect Chr(0) ' ʹ�ÿ��ַ��ƹ�����
    ws1.Unprotect Password:="123456" ' ���δ�������룬����ɾ�� Password ����
    ws2.Unprotect Password:="123456" ' ���δ�������룬����ɾ�� Password ����
    ws3.Unprotect Password:="123456" ' ���δ�������룬����ɾ�� Password ����

    ' ��ȡɸѡֵ�����������˵���A1��Ԫ��
    filterValue = ws1.Range("F1").Value

    ' ��ȡSheet1�����һ��
    lastRow1 = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row + 1

    ' ����Sheet2��C�У����ҷ�������������
    lastRow3 = ws3.Cells(ws3.Rows.Count, 8).End(xlUp).Row

    t = 11

    'ɾ����2��֮�����������
    ws1.Rows("10:" & ws1.Rows.Count).Delete Shift:=xlUp

    ' ���� Sheet1 �ĵ�һ������
    ws3.Rows(1).Copy

    ' ճ���� Sheet2 �ĵ�һ��
    ws1.Rows(t - 1).PasteSpecial Paste:=xlPasteAll

    For i = 2 To lastRow3 ' �ӵ�2�п�ʼ����
        If ws3.Cells(i, 3).Value = filterValue Then
            ' ���Ʒ������������е�Sheet1��ĩβ
            ws3.Rows(i).Copy Destination:=ws1.Rows(t)
            t = t + 1
        End If
    Next i
    If t = 11 Then
        MsgBox "�޽�����ʷ��"
    Else
        MsgBox "��ѯ�ɹ���"
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

Sub SerchImprotant()
    '
    ' ��ѯ�����ʲ� Macro
    '
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet
    Dim filterValue As String
    Dim lastRow1 As Long, lastRow2 As Long, lastRow3 As Long
    Dim i As Long, j As Long, t As Long

    ' ���ù�����
    Set ws1 = ThisWorkbook.Sheets("�������") ' ��ǰ������
    Set ws2 = ThisWorkbook.Sheets("�û�����") ' ����Դ������
    Set ws3 = ThisWorkbook.Sheets("�ʲ��嵥") ' ����Դ������

    ' ȡ�������������ƹ����룩
    'ws1.Unprotect Chr(0) ' ʹ�ÿ��ַ��ƹ�����
    ws1.Unprotect Password:="123456" ' ���δ�������룬����ɾ�� Password ����
    ws2.Unprotect Password:="123456" ' ���δ�������룬����ɾ�� Password ����
    ws3.Unprotect Password:="123456" ' ���δ�������룬����ɾ�� Password ����

    ' ��ȡɸѡֵ�����������˵���A1��Ԫ��
    filterValue = ws1.Range("F1").Value

    ' ��ȡSheet1�����һ��
    lastRow1 = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row + 1

    ' ����Sheet2��C�У����ҷ�������������
    lastRow3 = ws3.Cells(ws3.Rows.Count, 8).End(xlUp).Row

    lastRow2 = ws2.Cells(ws2.Rows.Count, 2).End(xlUp).Row
    t = 11

    'ɾ����2��֮�����������
    ws1.Rows("10:" & ws1.Rows.Count).Delete Shift:=xlUp

    ' ���� Sheet1 �ĵ�һ������
    ws3.Rows(1).Copy

    ' ճ���� Sheet2 �ĵ�һ��
    ws1.Rows(t - 1).PasteSpecial Paste:=xlPasteAll
    For i = 2 To lastRow3 ' �ӵ�2�п�ʼ����
        If Not IsEmpty(ws3.Cells(i, 9).Value) Then
            ' ���Ʒ������������е�Sheet1��ĩβ
            ws3.Rows(i).Copy Destination:=ws1.Rows(t)
            t = t + 1
        End If
    Next i
    If t = 11 Then
        MsgBox "�������ʲ���"
    Else
        MsgBox "��ѯ��ɣ�"
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
    'ws1.Rows(10 & ":" & t).Locked = False
    ' ������������ֹ�޸������ĵ�Ԫ��
    ws1.Protect Password:="123456", AllowFormattingCells:=True
    ws2.Protect Password:="123456", AllowFormattingCells:=True
    ws3.Protect Password:="123456", AllowFormattingCells:=True
    ThisWorkbook.Save '���浱ǰ������
End Sub

Sub SerchByType()

    Dim filterValue As String
    Dim LastDataRow As Long, LastAssetsRow As Long
    Dim matchedRange As Range, cell As Range

    Application.ScreenUpdating = False

    ' ��ȡɸѡֵ�����������˵���A1��Ԫ��
    filterValue = ManageSheet.Range(TYPE_FILTER_CELL).Value

    ' ȡ����������
    SheetUnLock ManageSheet

    ' ��ȡ�����
    LastDataRow = GetLastDataRow(ManageSheet)

    If (LastDataRow >= TARGET_ROW_START) Then
        ManageSheet.Rows(TARGET_ROW_START & ":" & LastDataRow).Delete Shift:=xlUp
    End If

    LastAssetsRow=GetLastDataRow(AssetsSheet)

    For Each cell In AssetsSheet.Range(AssetsSheet.Cells(2, TYPE_COLUMN), AssetsSheet.Cells(LastAssetsRow, TYPE_COLUMN))
        If cell.Value = filterValue Then
            If matchedRange Is Nothing Then
                Set matchedRange = cell.EntireRow
            Else
                Set matchedRange = Union(matchedRange, cell.EntireRow)
            End If
        End If
    Next cell

    ' ��������ƥ����
    If Not matchedRange Is Nothing Then
        matchedRange.Copy ManageSheet.Cells(TARGET_ROW_START, 1)
    Else
        MsgBox "δ�ҵ�ƥ����豸���ͣ�"
    End If

    Application.ScreenUpdating = True
    SheetLock ManageSheet
End Sub

