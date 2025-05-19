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
  ShowWindow
End Sub

Private Sub WorkbookInit()
  CurrentPath = ThisWorkbook.Path
  ' 隐藏工作表
  HideWindow
  ' 保护工作簿结构
  WorkbookLock
End Sub

Private Sub SheetsInit()

  Set ManageSheet = ThisWorkbook.Sheets(ManageSheetName)
  Set AssetsSheet = ThisWorkbook.Sheets(AssetsSheetName)
  Set LossHandleSheet = ThisWorkbook.Sheets(LossHandleSheetName)
  Set LDAssetsSheet = ThisWorkbook.Sheets(LDAssetsSheetName)
  Set UserDataSheet = ThisWorkbook.Sheets(UserDataSheetName)

  HideSheet ManageSheetName
  HideSheet AssetsSheetName
  HideSheet LossHandleSheetName
  HideSheet LDAssetsSheetName
  HideSheet UserDataSheetName

  SheetLock ManageSheet
  SheetLock AssetsSheet
  SheetLock LossHandleSheet
  SheetLock LDAssetsSheet
  SheetLock UserDataSheet

End Sub

Private Sub Login()
  UserID = Environ("COMPUTERNAME")
  If GetUserName(UserID) Then
    LogPrintf "用户 " & UserName & " 已登陆"
  Else
    LogPrintf "用户 " & UserID & " 未注册"
  End If
End Sub

Private Function GetUserName(searchID As String) As Boolean
  GetUserName = False

  Dim searchRange As Range
  Dim foundCell As Range
  With UserDataSheet
    Set searchRange = .Range("C1:C" & .Cells(.Rows.Count, 3).End(xlUp).Row)
    ' 精确查找（区分大小写）
    Set foundCell = searchRange.Find(What:=searchID, _
    LookIn:=xlValues, _
    LookAt:=xlWhole, _
    MatchCase:=True)

    If Not foundCell Is Nothing Then
      LogPrintf ("匹配项位于：" & foundCell.Address(0, 0))
      UserName = foundCell.Offset(0, -1).Value
      GetUserName = True
    End If

  End With

End Function



