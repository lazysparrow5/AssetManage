Attribute VB_Name = "查询"

Sub searchByName()
    '
    ' 查询 Macro
    '
    Dim filterValue As String

    ' 获取筛选值（假设下拉菜单在A1单元格）
    filterValue = ManageSheet.Range(NAME_FILTER_CELL).Value

    ' EditableClear

    ' ManageSheet.Range(NAME_FILTER_CELL).Value = filterValue

    TargetSearchCopy filterValue, USER_COLUMN, False

End Sub

Sub searchImprotant()
    '
    ' 查询入资资产 Macro
    '
    Dim filterValue As String

    filterValue = ""

    TargetSearchCopy filterValue, ASSETS_COLUMN, True

End Sub

Sub searchByType()

    Dim filterValue As String
    ' 获取筛选值（假设下拉菜单在A1单元格）
    filterValue = ManageSheet.Range(TYPE_FILTER_CELL).Value

    ' EditableClear

    ' ManageSheet.Range(TYPE_FILTER_CELL).Value = filterValue

    TargetSearchCopy filterValue, TYPE_COLUMN, False

End Sub

Public Sub TargetSearchCopy(filterValue As String, TargetColumn As Long, exclusion As Boolean)

    Dim LastDataRow As Long
    Dim matchedRange As Range

    Application.ScreenUpdating = False
    SheetUnLock ManageSheet
    ' 获取最低行
    LastDataRow = GetLastDataRow(ManageSheet)

    If (LastDataRow >= TARGET_ROW_START) Then
        ManageSheet.Rows(TARGET_ROW_START & ":" & LastDataRow).Delete Shift:=xlUp
    End If

    If exclusion = True Then
        ExcTargetSearch matchedRange, TargetColumn, filterValue
    Else
        TargetSearch matchedRange, TargetColumn, filterValue
    End If

    ' 批量复制匹配行
    If Not matchedRange Is Nothing Then
        matchedRange.Copy ManageSheet.Cells(TARGET_ROW_START, 1)
    Else
        MsgBox "未找到匹配的设备！"
    End If

    SheetLock ManageSheet
    Application.ScreenUpdating = True
End Sub

Public Sub TargetSearch(matchedRange As Range, TargetColumn As Long, filterValue As String)
    
    Dim cell As Range

    For Each cell In AssetsSheet.Range(AssetsSheet.Cells(2, TargetColumn), AssetsSheet.Cells(AssetsIndexMax, TargetColumn))
        If cell.Value = filterValue Then
            If matchedRange Is Nothing Then
                Set matchedRange = cell.EntireRow
            Else
                Set matchedRange = Union(matchedRange, cell.EntireRow)
            End If
        End If
    Next cell

End Sub

Public Sub ExcTargetSearch(matchedRange As Range, TargetColumn As Long, filterValue As String)
    
    Dim cell As Range

    For Each cell In AssetsSheet.Range(AssetsSheet.Cells(2, TargetColumn), AssetsSheet.Cells(AssetsIndexMax, TargetColumn))
        If cell.Value <> filterValue Then
            If matchedRange Is Nothing Then
                Set matchedRange = cell.EntireRow
            Else
                Set matchedRange = Union(matchedRange, cell.EntireRow)
            End If
        End If
    Next cell

End Sub

