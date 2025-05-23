Attribute VB_Name = "核查"
Sub AssetsCheck()
    '
    ' 生成盘点表 Macro
    '
    Dim filterValue As String
    ' 获取筛选值（假设下拉菜单在A1单元格）
    filterValue = ""

    LogPrintf "生成盘点表 (" & UserName & ")", Log_User

    TargetSearchCopy filterValue, USER_COLUMN, True

    Dim selectedRange As Range

    Set selectedRange = Selection

    If selectedRange Is Nothing Or selectedRange.Columns.Count < USER_COLUMN Then
        Exit Sub
    End If

    selectedRange.Sort Key1:=selectedRange.Columns(USER_COLUMN), Order1:=xlAscending, Header:=xlNo 

    LastSearch = CheckSearch

End Sub
