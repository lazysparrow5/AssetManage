Attribute VB_Name = "ȫ�ֱ���"

'�û�����
Public Const AdminID As String = "XING"
Public Const AdminPwd As String = "123456"

Public UserName As String '�û���'
Public UserID As String '�û�ID'

'����������
Public Const ManageBookName As String = "Assets.xlsm"
Public Const DataBookName As String = "DataBase.xlsm"

Public Const StartSheetName As String = "��ʼ����"
Public Const ManageSheetName As String = "�������"

Public Const AssetsSheetName As String = "�ʲ��嵥"
Public Const UserDataSheetName As String = "�û�����"

Public Const TYPE_FILTER_CELL As String = "B1"        ' ��������λ��
Public Const NAME_FILTER_CELL As String = "F1"     ' ��������λ��
Public Const INDEX_BORROW_CELL As String = "B7"     ' �����������λ��
Public Const USER_BORROW_CELL As String = "C7"      ' ��������������λ��
Public Const BRIEF_BORROW_CELL As String = "D7"     ' ������Ʒ�������λ��
Public Const INDEX_RETURN_CELL As String = "G7"     ' �黹�������λ��
Public Const POSITION_RETURN_CELL As String = "H7"      ' �黹��ַ����λ��
Public Const BRIEF_RETURN_CELL As String = "I7"     ' �黹��Ʒ�������λ��

' �ʲ��������
Public Const TARGET_ROW_START As Long = 11       ' Ŀ����ʼ��
Public Const INDEX_COLUMN As Long = 1           ' �豸���
Public Const TYPE_COLUMN As Long = 8             ' �豸�����к�
Public Const USER_COLUMN As Long = 3              ' �豸�����к�
Public Const ASSETS_COLUMN As Long = 9             '�ʲ�����к�

' Public Const LDAssetsSheetName As String = "�����ʲ�"
' Public Const LossHandleSheetName As String = "��ʧ�⳥"

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






