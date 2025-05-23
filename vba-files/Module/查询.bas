Attribute VB_Name = "查询"

Public Sub searchByName()
    '
    ' 查询 Macro
    '
    Dim filterValue As String

    ' 获取筛选值（假设下拉菜单在A1单元格）
    filterValue = ManageSheet.Range(NAME_FILTER_CELL).Value

    LogPrintf "通过名字[" & filterValue & "]查询 (" & UserName & ")", Log_User

    ' EditableClear

    ' ManageSheet.Range(NAME_FILTER_CELL).Value = filterValue

    TargetSearchCopy filterValue, USER_COLUMN, False

    LastSearch = NameSearch

End Sub

Public Sub searchImprotant()
    '
    ' 查询入资资产 Macro
    '
    Dim filterValue As String

    filterValue = ""

    LogPrintf "查询入资资产 (" & UserName & ")" , Log_User

    TargetSearchCopy filterValue, SGMW_COLUMN, True

    LastSearch = ImportantSearch

End Sub

Public Sub searchByType()

    Dim filterValue As String
    ' 获取筛选值（假设下拉菜单在A1单元格）
    filterValue = ManageSheet.Range(TYPE_FILTER_CELL).Value

    LogPrintf "通过类型{" & filterValue & "}查询 (" & UserName & ")", Log_User

    ' EditableClear

    ' ManageSheet.Range(TYPE_FILTER_CELL).Value = filterValue

    TargetSearchCopy filterValue, TYPE_COLUMN, False

    LastSearch = TypeSearch

End Sub

Public Sub ReSearch()

    Select Case LastSearch 
     Case TypeSearch 
        searchByType  
     Case NameSearch 
        searchByName  
     Case ImportantSearch 
        searchImprotant  
     Case CheckSearch 
        AssetsCheck
    End Select

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
        ' MsgBox "未找到匹配的设备！"
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

Public Function NameIsExsits(sName As String) As Boolean
  NameIsExsits = False

  Dim searchRange As Range
  Dim foundCell As Range
  With UserDataSheet
    Set searchRange = .Range("A1:A" & .Cells(.Rows.Count, 1).End(xlUp).Row)
    ' 精确查找（区分大小写）
    Set foundCell = searchRange.Find(What:=sName, _
    LookIn:=xlValues, _
    LookAt:=xlWhole, _
    MatchCase:=True)

    If Not foundCell Is Nothing Then
      ' LogPrintf ("匹配项位于：" & foundCell.Address(0, 0))
      NameIsExsits = True
    End If

  End With

End Function

