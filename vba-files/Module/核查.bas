Attribute VB_Name = "�˲�"
Sub AssetsCheck()
    '
    ' �����̵�� Macro
    '
    Dim filterValue As String
    ' ��ȡɸѡֵ�����������˵���A1��Ԫ��
    filterValue = ""

    LogPrintf "�����̵�� (" & UserName & ")", Log_User

    TargetSearchCopy filterValue, USER_COLUMN, True

    Dim selectedRange As Range

    Set selectedRange = Selection

    If selectedRange Is Nothing Or selectedRange.Columns.Count < USER_COLUMN Then
        Exit Sub
    End If

    selectedRange.Sort Key1:=selectedRange.Columns(USER_COLUMN), Order1:=xlAscending, Header:=xlNo 

    LastSearch = CheckSearch

End Sub
