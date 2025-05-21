Attribute VB_Name = "全局变量"

'用户数据
Public Const AdminID As String = "XING"
Public Const AdminPwd As String = "123456"

Public UserName As String '用户名'
Public UserID As String '用户ID'

'管理表格数据
Public Const ManageBookName As String = "Assets.xlsm"
Public Const DataBookName As String = "DataBase.xlsm"

Public Const StartSheetName As String = "开始界面"
Public Const ManageSheetName As String = "管理界面"

Public Const AssetsSheetName As String = "资产清单"
Public Const UserDataSheetName As String = "用户数据"

' Public Const LDAssetsSheetName As String = "柳东资产"
' Public Const LossHandleSheetName As String = "遗失赔偿"

Public CurrentPath As String

Public ManageBook As Workbook
Public DataBook As Workbook

Public StartSheet As Worksheet
Public ManageSheet As Worksheet
Public AssetsSheet As Worksheet
Public UserDataSheet As Worksheet
' Public LossHandleSheet As Worksheet
' Public LDAssetsSheet As Worksheet




