Attribute VB_Name = "ȫ�ֱ���"

'�û�����
Public Const AdminID As String = "XING"
Public UserName As String '�û���'
Public UserID As String '�û�ID'

'�������
Public Const BookName As String = "Assets.xlsm"
Public Const ManageSheetName As String = "�������"
Public Const AssetsSheetName As String = "�ʲ��嵥"
Public Const LossHandleSheetName As String = "��ʧ�⳥"
Public Const LDAssetsSheetName As String = "�����ʲ�"
Public Const UserDataSheetName As String = "�û�����"

Public Const DefaultSheetName As String = "�������"

Public ManageSheet As Worksheet
Public AssetsSheet As Worksheet
Public LossHandleSheet As Worksheet
Public LDAssetsSheet As Worksheet
Public UserDataSheet As Worksheet
Public DefaultSheet As Worksheet

Public CurrentPath As String
