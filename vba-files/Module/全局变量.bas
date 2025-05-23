Attribute VB_Name = "全局变量"

Public Enum SearchType
TypeSearch = 1 
NameSearch
ImportantSearch
CheckSearch
End Enum

Public Enum LogType
Log_Debug = 1
Log_User 
Log_Err
Log_Assets
End Enum

'用户数据
Public Const AdminID As String = "陈聪"
Public Const AdminPwd As String = "123456"

Public UserName As String '用户名'
Public UserID As String '用户ID'

Public LastSearch As SearchType

'管理表格数据
Public Const ManageBookName As String = "Assets.xlsm"
Public Const DataBookName As String = "DataBase.xlsm"

Public Const StartSheetName As String = "开始界面"
Public Const ManageSheetName As String = "管理界面"

Public Const AssetsSheetName As String = "资产清单"
Public Const UserDataSheetName As String = "用户数据"

Public Const TYPE_FILTER_CELL As String = "B1"        ' 类型输入位置
Public Const NAME_FILTER_CELL As String = "F1"     ' 名称输入位置
Public Const INDEX_BORROW_CELL As String = "B7"     ' 借用序号输入位置
Public Const USER_BORROW_CELL As String = "C7"      ' 借用人名称输入位置
Public Const BRIEF_BORROW_CELL As String = "D7"     ' 借用物品简介输入位置
Public Const INDEX_RETURN_CELL As String = "G7"     ' 归还序号输入位置
Public Const USER_RETURN_CELL As String = "H7"      ' 归还地址输入位置
Public Const BRIEF_RETURN_CELL As String = "I7"     ' 归还物品简介输入位置

' 资产表格数据
Public Const TARGET_ROW_START As Long = 11       ' 目标起始行
Public Const INDEX_COLUMN As Long = 1           ' 设备序号
Public Const NAME_COLUMN As Long =2                 '设备名称列号
Public Const USER_COLUMN As Long = 3              ' 用户名称列号
Public Const LOCATION_COLUMN As Long = 4          ' 存储地址列号
Public Const BRIEF_COLUMN As Long = 5              ' 简介列号
Public Const REVISER_COLUMN As Long = 6             ' 修改者列号
Public Const WHOADD_COLUMN As Long = 7              ' 入库人
Public Const TYPE_COLUMN As Long = 8             ' 设备类型列号
Public Const SGMW_COLUMN As Long = 9             '资产编号列号
Public Const ORIGIN_COLUMN As Long = 10              ' 原厂编号列号
Public Const TIME_COLUMN As Long = 11              ' 修改时间列号
Public Const CHECK_COLUMN As Long = 12              ' 盘点结果列号
Public Const MAX_COLUMN As Long = 13                '最大列号

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

Public AssetsIndexMax As Long






