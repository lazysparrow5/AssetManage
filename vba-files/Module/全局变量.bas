Attribute VB_Name = "ȫ�ֱ���"

Public Enum SearchType
TypeSearch = 1 
NameSearch
ImportantSearch
CheckSearch
End Enum

'�û�����
Public Const AdminID As String = "�´�"
Public Const AdminPwd As String = "123456"

Public UserName As String '�û���'
Public UserID As String '�û�ID'

Public LastSearch As SearchType

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
Public Const USER_RETURN_CELL As String = "H7"      ' �黹��ַ����λ��
Public Const BRIEF_RETURN_CELL As String = "I7"     ' �黹��Ʒ�������λ��

' �ʲ��������
Public Const TARGET_ROW_START As Long = 11       ' Ŀ����ʼ��
Public Const INDEX_COLUMN As Long = 1           ' �豸���
Public Const NAME_COLUMN As Long =2                 '�豸�����к�
Public Const USER_COLUMN As Long = 3              ' �û������к�
Public Const LOCATION_COLUMN As Long = 4          ' �洢��ַ�к�
Public Const BRIEF_COLUMN As Long = 5              ' ����к�
Public Const REVISER_COLUMN As Long = 6             ' �޸����к�
Public Const WHOADD_COLUMN As Long = 7              ' �����
Public Const TYPE_COLUMN As Long = 8             ' �豸�����к�
Public Const SGMW_COLUMN As Long = 9             '�ʲ�����к�
Public Const ORIGIN_COLUMN As Long = 10              ' ԭ������к�
Public Const TIME_COLUMN As Long = 11              ' �޸�ʱ���к�
Public Const CHECK_COLUMN As Long = 12              ' �̵����к�
Public Const MAX_COLUMN As Long = 13                '����к�

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






