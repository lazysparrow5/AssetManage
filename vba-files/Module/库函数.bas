Attribute VB_Name = "�⺯��"

' �������
Sub SheetLock(ws As Worksheet)
    ws.Protect Password:=AdminPwd, AllowFormattingCells:=False, AllowSorting:=False, AllowFiltering:=False  ' �����������ʽ ���������� ������ɸѡ
End Sub

'������
Public Sub SheetUnLock(ws As Worksheet)
    ws.Unprotect Password:=AdminPwd
End Sub

'����������
Public Sub WorkbookLock()
    ThisWorkbook.Protect Password:=AdminPwd, Structure:=True, Windows:=True      '��ֹ��ɾ�������ƶ�ͼ�� ��ֹ�������ڲ���
End Sub

'����������
Public Sub WorkbookUnLock()
    ThisWorkbook.Unprotect Password:=AdminPwd
End Sub

Public Sub EditableUnLock()
    ' �Ƚ����������������ѱ�����
    SheetUnLock ManageSheet
    With ManageSheet
        ' �������÷������������
        .Range("B1,F1").Locked = False        ' ����������Ԫ��
        .Range("B4:H4").Locked = False        ' ��4����������
        .Range("B7:D7,G7:I7").Locked = False  ' ��7��������������
    End With
    ' ���±�������������������Ԫ������״̬��
    SheetLock ManageSheet
End Sub

Public Sub EditableClear()
    With ManageSheet
        ' �����������Ԫ��
        .Range("B1,F1").ClearContents
        ' ����������� B4:H4��7����Ԫ��
        .Range("B4:H4").ClearContents
        ' ���������������B7:D7 �� G7:I7��
        .Range("B7:D7,G7:I7").ClearContents
    End With
End Sub

Public Sub ShowWindow()
    ThisWorkbook.Windows(BookName).Visible = True ' ��ʾ��ǰ�������Ĵ���
End Sub

Public Sub HideWindow()
    ThisWorkbook.Windows(BookName).Visible = False ' ��ʾ��ǰ�������Ĵ���
End Sub

Public Sub ShowSheet(SheetName As String)
    WorkbookUnLock
    On Error Resume Next
    Sheets(SheetName).Visible = xlSheetVisible
    On Error GoTo 0
    WorkbookLock
End Sub

Public Sub HideSheet(SheetName As String)
    WorkbookUnLock
    On Error Resume Next
    Sheets(SheetName).Visible = xlSheetVeryHidden
    On Error GoTo 0
    WorkbookLock
End Sub

' ���ϴ���
Sub ErrHandle()


End Sub

Sub LogPrintf(logMessage As String)
    Dim LogPath As String
    LogPath = CurrentPath & "\log.txt" ' ��ǰLog����·��

    ' �����ļ�ϵͳ����
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' ���ļ���׷��ģʽ���������򴴽���
    Dim logFile As Object
    Set logFile = fso.OpenTextFile(LogPath, 8, True) ' 8 = ׷��ģʽ

    ' д���ʱ�������־
    logFile.WriteLine Format(Now, "yyyy-mm-dd hh:mm:ss") & " - " & logMessage
    logFile.Close

End Sub
