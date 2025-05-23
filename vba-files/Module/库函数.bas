Attribute VB_Name = "库函数"

' 表格上锁
Sub SheetLock(ws As Worksheet)
    ws.Protect Password:=AdminPwd, AllowFormattingCells:=False, AllowSorting:=False, AllowFiltering:=False  ' 不允许调整格式 不允许排序 不允许筛选
End Sub

'表格解锁
Public Sub SheetUnLock(ws As Worksheet)
    ws.Unprotect Password:=AdminPwd
End Sub

'工作簿上锁
Public Sub WorkbookLock(wb As Workbook)
    wb.Protect Password:=AdminPwd, Structure:=True, Windows:=True      '禁止增删工作表、移动图表 禁止调整窗口布局
End Sub

'工作簿解锁
Public Sub WorkbookUnLock(wb As Workbook)
    wb.Unprotect Password:=AdminPwd
End Sub

Public Sub EditableUnLock()
    ' 先解除工作表保护（如果已保护）
    SheetUnLock ManageSheet
    With ManageSheet
        ' 批量设置非连续区域解锁
        .Range("B1,F1").Locked = False        ' 顶部独立单元格
        .Range("B4:H4").Locked = False        ' 第4行连续区域
        .Range("B7:D7,G7:I7").Locked = False  ' 第7行两个连续区域
    End With
    ' 重新保护工作表（保持其他单元格锁定状态）
    SheetLock ManageSheet
End Sub

Public Sub EditableClear()
    With ManageSheet
        ' 清除非连续单元格
        .Range("B1,F1").ClearContents
        ' 清空连续区域 B4:H4（7个单元格）
        .Range("B4:H4").ClearContents
        ' 清空两段连续区域（B7:D7 和 G7:I7）
        .Range("B7:D7,G7:I7").ClearContents
    End With
End Sub

Public Sub ShowWindow(wb As Workbook)
    wb.Windows(1).Visible = True ' 显示当前工作簿的窗口
End Sub

Public Sub HideWindow(wb As Workbook)
    wb.Windows(1).Visible = False ' 隐藏当前工作簿的窗口
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


' 故障处理
Sub ErrHandle()


End Sub

Public Sub LogPrintf(logMessage As String, Optional sLogtype As LogType = Log_Debug)
    
    Dim LogPath As String
    
    Select Case sLogtype
     Case Log_Debug
        LogPath = CurrentPath & "\log_Debug.txt" ' 当前Log所在路径
     Case Log_User
        LogPath = CurrentPath & "\log_User.txt"
     Case Log_Err
        LogPath = CurrentPath & "\log_Err.txt"
     Case Log_Assets
        LogPath = CurrentPath & "\log_Assets.txt"
    End Select
    
    ' 创建文件系统对象
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' 打开文件（追加模式，不存在则创建）
    Dim logFile As Object
    Set logFile = fso.OpenTextFile(LogPath, 8, True) ' 8 = 追加模式

    ' 写入带时间戳的日志
    logFile.WriteLine Format(Now, "yyyy-mm-dd hh:mm:ss") & " - " & logMessage
    logFile.Close

End Sub
