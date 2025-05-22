Attribute VB_Name = "��ѯ"

Public Sub searchByName()
    '
    ' ��ѯ Macro
    '
    Dim filterValue As String

    ' ��ȡɸѡֵ�����������˵���A1��Ԫ��
    filterValue = ManageSheet.Range(NAME_FILTER_CELL).Value

    ' EditableClear

    ' ManageSheet.Range(NAME_FILTER_CELL).Value = filterValue

    TargetSearchCopy filterValue, USER_COLUMN, False

    LastSearch = NameSearch

End Sub

Public Sub searchImprotant()
    '
    ' ��ѯ�����ʲ� Macro
    '
    Dim filterValue As String

    filterValue = ""

    TargetSearchCopy filterValue, SGMW_COLUMN, True

    LastSearch = ImportantSearch

End Sub

Public Sub searchByType()

    Dim filterValue As String
    ' ��ȡɸѡֵ�����������˵���A1��Ԫ��
    filterValue = ManageSheet.Range(TYPE_FILTER_CELL).Value

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
    ' ��ȡ�����
    LastDataRow = GetLastDataRow(ManageSheet)

    If (LastDataRow >= TARGET_ROW_START) Then
        ManageSheet.Rows(TARGET_ROW_START & ":" & LastDataRow).Delete Shift:=xlUp
    End If

    If exclusion = True Then
        ExcTargetSearch matchedRange, TargetColumn, filterValue
    Else
        TargetSearch matchedRange, TargetColumn, filterValue
    End If

    ' ��������ƥ����
    If Not matchedRange Is Nothing Then
        matchedRange.Copy ManageSheet.Cells(TARGET_ROW_START, 1)
    Else
        ' MsgBox "δ�ҵ�ƥ����豸��"
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

