Attribute VB_Name = "�⺯��"

' �������
Sub SheetLock()


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
