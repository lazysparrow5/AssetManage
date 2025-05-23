Attribute VB_Name = "借用"

Private IndexBorrowValue As String
Private UserBorrowValue As String
Private BriefBorrowValue As String

Private IndexReturnValue As String
Private UserReturnValue As String
Private BriefReturnValue As String

Sub UserBorrow()
    '
    ' 点击借用
    '
    Dim DataBaseSheet As Worksheet
    Dim LastRowsData As Range, CurrRowsData As Range

    IndexBorrowValue = ManageSheet.Range(INDEX_BORROW_CELL).Value
    UserBorrowValue = ManageSheet.Range(USER_BORROW_CELL).Value
    BriefBorrowValue = ManageSheet.Range(BRIEF_BORROW_CELL).Value

    If Not BorrowInputValid Then
     Exit Sub  
    End If

    Set LastRowsData = AssetsSheet.Rows(IndexBorrowValue + 1)

    DataBaseOpen
    DataUpdate

    Set DataBaseSheet = DataBook.Worksheets(AssetsSheetName)
    Set CurrRowsData = DataBaseSheet.Rows(IndexBorrowValue + 1)

    If Not RowsDataIsSame(LastRowsData,CurrRowsData) Then
        MsgBox "该设备信息已被修改，请重试"
        DataBaseClose
     Exit Sub  
    End If 

    SheetUnlock DataBaseSheet

    With CurrRowsData
        .Cells(USER_COLUMN).Value = UserBorrowValue
        .Cells(REVISER_COLUMN).Value = UserName
        .Cells(TIME_COLUMN).Value = Date
        If Trim(BriefBorrowValue) <> "" Then
            .Cells(BRIEF_COLUMN) = BriefBorrowValue
        End If
    End With

    SheetLock DataBaseSheet

    SheetUnLock AssetsSheet
    CurrRowsData.Copy Destination:=LastRowsData
    SheetLock AssetsSheet

    ManageSheet.Range(INDEX_BORROW_CELL).ClearContents
    ManageSheet.Range(USER_BORROW_CELL).ClearContents
    ManageSheet.Range(BRIEF_BORROW_CELL).ClearContents

    DataBaseClose

    ReSearch

End Sub

Sub UserReturn()
    '
    ' 点击归还 Macro
    '
    If UserName <> AdminID Then
        MsgBox "归还处于维护中，请联系管理员归还"
        Exit Sub
    End If

    Dim DataBaseSheet As Worksheet
    Dim LastRowsData As Range, CurrRowsData As Range

    IndexReturnValue = ManageSheet.Range(INDEX_RETURN_CELL).Value
    UserReturnValue = ManageSheet.Range(USER_RETURN_CELL).Value
    BriefReturnValue = ManageSheet.Range(BRIEF_RETURN_CELL).Value

    If Not ReturnInputValid Then
     Exit Sub  
    End If

    Set LastRowsData = AssetsSheet.Rows(IndexReturnValue + 1)

    DataBaseOpen
    DataUpdate

    Set DataBaseSheet = DataBook.Worksheets(AssetsSheetName)
    Set CurrRowsData = DataBaseSheet.Rows(IndexBorrowValue + 1)

    If Not RowsDataIsSame(LastRowsData,CurrRowsData) Then
        MsgBox "该设备信息已被修改，请重试"
        DataBaseClose
     Exit Sub  
    End If 

    SheetUnlock DataBaseSheet

    With CurrRowsData
        .Cells(USER_COLUMN).ClearContents
        .Cells(LOCATION_COLUMN).Value = UserReturnValue
        .Cells(TIME_COLUMN).Value = Date
        If Trim(BriefReturnValue) <> "" Then
            .Cells(BRIEF_COLUMN) = BriefReturnValue
        End If
    End With

    SheetLock DataBaseSheet

    SheetUnLock AssetsSheet
    CurrRowsData.Copy Destination:=LastRowsData
    SheetLock AssetsSheet

    ManageSheet.Range(INDEX_RETURN_CELL).ClearContents
    ManageSheet.Range(USER_RETURN_CELL).ClearContents
    ManageSheet.Range(BRIEF_RETURN_CELL).ClearContents

    DataBaseClose

    ReSearch

End Sub

Private Function BorrowInputValid() As Boolean

    BorrowInputValid = False

    Select Case True
     Case Trim(IndexBorrowValue) = ""
        MsgBox "序号不能为空"
     Case Not IsNumeric(IndexBorrowValue)
        MsgBox "序号必须为数字"
     Case IndexBorrowValue <> Fix(IndexBorrowValue)
        MsgBox "序号必须为整数"
     Case IndexBorrowValue <= 1 & IndexBorrowValue >= AssetsIndexMax
        MsgBox "序号不在范围内"
     Case Trim(UserBorrowValue) = ""
        MsgBox "借用人不能为空"
     Case Else
        BorrowInputValid = True
    End Select

End Function

Private Function ReturnInputValid() As Boolean
    
    ReturnInputValid = False

    Select Case True
     Case Trim(IndexBorrowValue) = ""
        MsgBox "序号不能为空"
     Case Not IsNumeric(IndexBorrowValue)
        MsgBox "序号必须为数字"
     Case IndexBorrowValue <> Fix(IndexBorrowValue)
        MsgBox "序号必须为整数"
     Case IndexBorrowValue <= 1 & IndexBorrowValue >= AssetsIndexMax
        MsgBox "序号不在范围内"
     Case Trim(UserReturnValue) = ""
        MsgBox "归还地址不能为空"
     Case Else
        ReturnInputValid = True
    End Select

End Function


