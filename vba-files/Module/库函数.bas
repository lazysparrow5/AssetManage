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
Public Sub WorkbookLock(wb As Workbook)
    wb.Protect Password:=AdminPwd, Structure:=True, Windows:=True      '��ֹ��ɾ�������ƶ�ͼ�� ��ֹ�������ڲ���
End Sub

'����������
Public Sub WorkbookUnLock(wb As Workbook)
    wb.Unprotect Password:=AdminPwd
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

Public Sub ShowWindow(wb As Workbook)
    wb.Windows(1).Visible = True ' ��ʾ��ǰ�������Ĵ���
End Sub

Public Sub HideWindow(wb As Workbook)
    wb.Windows(1).Visible = False ' ���ص�ǰ�������Ĵ���
End Sub

Public Sub ShowSheet(ws As Worksheet)
    ws.Visible = xlSheetVisible
End Sub

Public Sub HideSheet(ws As Worksheet)
    ws.Visible = xlSheetVeryHidden
End Sub

Private Sub SetSheetHandle()
    If SheetExists(ManageBook,StartSheetName) Then
        Set StartSheet = ManageBook.Sheets(StartSheetName)
    End If
    If SheetExists(ManageBook,ManageSheetName) Then
        Set ManageSheet = ManageBook.Sheets(ManageSheetName)
    End If
    If SheetExists(ManageBook,AssetsSheetName) Then
        Set AssetsSheet = ManageBook.Sheets(AssetsSheetName)
    End If
    If SheetExists(ManageBook,UserDataSheetName) Then
        Set UserDataSheet = ManageBook.Sheets(UserDataSheetName)
    End If
End Sub

Public Sub GetDataBase()

    Application.DisplayAlerts = False
    If SheetExists(ManageBook,AssetsSheetName) Then
        ManageBook.Sheets(AssetsSheetName).Delete
    End If
    If SheetExists(ManageBook,UserDataSheetName) Then
        ManageBook.Sheets(UserDataSheetName).Delete
    End If
    Application.DisplayAlerts = True
    
    DataBook.Sheets(AssetsSheetName).Copy After:=ManageBook.Sheets(ManageBook.Sheets.Count)
    DataBook.Sheets(UserDataSheetName).Copy After:=ManageBook.Sheets(ManageBook.Sheets.Count)

End Sub

Public Sub DataBaseOpen()
    Set DataBook = Workbooks.Open(CurrentPath & "\" & DataBookName)
    HideWindow DataBook
End Sub

Public Sub DataBaseClose()
    DataBook.Close SaveChanges:=True
    Set DataBook = Nothing
End Sub

Public Sub DataUpdate()

    WorkbookUnLock ManageBook
    GetDataBase
    SetSheetHandle
    HideSheet AssetsSheet
    HideSheet UserDataSheet
    WorkbookLock ManageBook
    AssetsIndexMax = GetLastDataRow(AssetsSheet)

End Sub

Public Function GetLastDataRow(ws As Worksheet, Optional Byval columnNumber As Long = 1) As Long
    GetLastDataRow = ws.Cells(ws.Rows.Count, columnNumber).End(xlUp).Row
End Function

Public Function RowsDataIsSame(row1data As Variant, row2data As Variant) As Boolean

    RowsDataIsSame = (Join(Application.Index(row1data, 1, 0),"|") = Join(Application.Index(row2data, 1, 0), "|")) 

End Function

Public Function SheetExists(wb As Workbook, sName As String) As Boolean
    SheetExists = False
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(sName)
    On Error GoTo 0
    If Not ws Is Nothing Then
        SheetExists = True
    End If
End Function


' ���ϴ���
Sub ErrHandle()


End Sub

Public Sub LogPrintf(logMessage As String, Optional sLogtype As LogType = Log_Debug)
    
    Dim LogPath As String
    
    Select Case sLogtype
     Case Log_Debug
        LogPath = CurrentPath & "\log_Debug.txt" ' ��ǰLog����·��
     Case Log_User
        LogPath = CurrentPath & "\log_User.txt"
     Case Log_Err
        LogPath = CurrentPath & "\log_Err.txt"
     Case Log_Assets
        LogPath = CurrentPath & "\log_Assets.txt"
    End Select
    
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
