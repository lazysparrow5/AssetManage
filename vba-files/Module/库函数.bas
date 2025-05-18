Attribute VB_Name = "库函数"

' 表格上锁
Sub SheetLock()


End Sub

' 故障处理
Sub ErrHandle()


End Sub

Sub LogPrintf(logMessage As String)
    Dim LogPath As String
    LogPath = CurrentPath & "\log.txt" ' 当前Log所在路径

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
