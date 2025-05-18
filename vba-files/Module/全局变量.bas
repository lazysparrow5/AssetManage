Attribute VB_Name = "全局变量"

'用户数据
Public Const AdminID As String = "XING"
Public UserName As String '用户名'
Public UserID As String '用户ID'

'表格数据
Public Const BookName As String = "Assets.xlsm"
Public Const ManageSheetName As String = "管理界面"
Public Const AssetsSheetName As String = "资产清单"
Public Const LossHandleSheetName As String = "遗失赔偿"
Public Const LDAssetsSheetName As String = "柳东资产"
Public Const UserDataSheetName As String = "用户数据"

Public Const DefaultSheetName As String = "管理界面"

Public ManageSheet As Worksheet
Public AssetsSheet As Worksheet
Public LossHandleSheet As Worksheet
Public LDAssetsSheet As Worksheet
Public UserDataSheet As Worksheet
Public DefaultSheet As Worksheet

Public CurrentPath As String
