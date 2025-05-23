VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
  ' 这里是打开工作簿时要执行的代码
  WorkbookInit
  SheetsInit
  LogPrintf "欢迎使用资产管理系统"
  Login
  Application.Visible = True
End Sub

Private Sub WorkbookInit()
  CurrentPath = ThisWorkbook.Path
  Set ManageBook = Workbooks(ManageBookName)
  ' 隐藏工作簿
  ShowWindow ManageBook
  Application.Visible = False
  ' 保护工作簿结构
End Sub

Private Sub SheetsInit()

  DataInit

  ShowNoMacroWindow

  SheetControl

End Sub

Private Sub Login()
  UserID = Environ("COMPUTERNAME")
  If GetUserName(UserID) Then
    LogPrintf "用户 " & UserName & " 已登陆"
    WorkbookUnLock ManageBook
    ManageSheet.Range(NAME_FILTER_CELL).Value = UserName
    SearchByName
    ShowUserWindow
    WorkbookLock ManageBook
  Else
    LogPrintf "用户 " & UserID & " 未注册"
    ' MsgBox "未查询到用户，请联系管理员进行登记"
  End If
End Sub

Private Sub SheetControl()
  SheetLock StartSheet
  SheetLock ManageSheet
  SheetLock AssetsSheet
  SheetLock UserDataSheet
End Sub

Private Sub ShowUserWindow()
  WorkbookUnLock ManageBook
  ShowSheet ManageSheet

  HideSheet StartSheet
  HideSheet AssetsSheet
  HideSheet UserDataSheet
  WorkbookLock ManageBook
End Sub

Private Sub ShowAdminWindow()

  ShowWindow DataBook

  WorkbookUnLock ManageBook
  ShowSheet ManageSheet

  ShowSheet StartSheet
  ShowSheet AssetsSheet
  ShowSheet UserDataSheet
  WorkbookLock ManageBook

End Sub

Private  Sub ShowNoMacroWindow()
  WorkbookUnLock ManageBook
  ShowSheet StartSheet

  HideSheet ManageSheet
  HideSheet AssetsSheet
  HideSheet UserDataSheet
  WorkbookLock ManageBook
End Sub

Private Sub DataInit()
  DataBaseOpen
  DataUpdate
  DataBaseClose
End Sub

Private Function GetUserName(searchID As String) As Boolean
  GetUserName = False

  Dim searchRange As Range
  Dim foundCell As Range
  With UserDataSheet
    Set searchRange = .Range("B1:B" & .Cells(.Rows.Count, 2).End(xlUp).Row)
    ' 精确查找（区分大小写）
    Set foundCell = searchRange.Find(What:=searchID, _
    LookIn:=xlValues, _
    LookAt:=xlWhole, _
    MatchCase:=True)

    If Not foundCell Is Nothing Then
      ' LogPrintf ("匹配项位于：" & foundCell.Address(0, 0))
      UserName = foundCell.Offset(0, -1).Value
      GetUserName = True
    End If

  End With

End Function



