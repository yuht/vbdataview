VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   8805
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":059A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7646
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Download by http://www.NewXing.com

'���� ImageList1 �ؼ� �Ҽ� -> ���� �������ͼ��

Private Sub Form_Load()
    
    ListView1.SmallIcons = ImageList1.Object                                    '��ImageList1 �ؼ��󶨵� ListView1�б�ؼ������� ������Ҫ��ImageList1�е�ͼ��
    
    ListView1.ListItems.Clear                                                   '����б�
    ListView1.ColumnHeaders.Clear                                               '����б�ͷ
    ListView1.View = lvwReport                                                  '�����б���ʾ��ʽ
    ListView1.GridLines = True                                                  '��ʾ������
    ListView1.LabelEdit = lvwManual                                             '��ֹ��ǩ�༭
    ListView1.FullRowSelect = True                                              'ѡ������
    
    ListView1.ColumnHeaders.Add , , "ID", 500                                   '���б����������
    ListView1.ColumnHeaders.Add , , "���� IP", 1500
    ListView1.ColumnHeaders.Add , , "���ض˿�", 1200
    ListView1.ColumnHeaders.Add , , "Э��", 550
    ListView1.ColumnHeaders.Add , , "Զ�� IP", 1500
    ListView1.ColumnHeaders.Add , , "Զ�̶˿�", 900
    ListView1.ColumnHeaders.Add , , "��ǰ״̬", 900
    ListView1.ColumnHeaders.Add , , "����ʱ��", 900
    '-------------------------------------------------------
    Dim X
    X = ListView1.ListItems.Count + 1
    ListView1.ListItems.Add , , X
    ListView1.ListItems(X).SubItems(1) = "00:00:00"
    ListView1.ListItems(X).SubItems(2) = "2008-01-01"
    ListView1.ListItems(X).SubItems(3) = "(��)"
    '-------------------------------------------------------
    ListView1.ListItems.Clear                                                   '����б�
    ListView1.ListItems.Add , , "1", , 1                                        '���ͼ��  �����Ǹ�1��ImageList1�ؼ��е�ͼ��������
    ListView1.ListItems(1).SubItems(1) = "00:00:00"
    ListView1.ListItems(1).SubItems(2) = "2008-01-01"
    ListView1.ListItems(1).SubItems(3) = "(��)"
    ListView1.ListItems(1).SubItems(4) = "127.0.0.1"
    
    ListView1.ListItems.Add , , "2", , 2                                        '���ͼ��  �����Ǹ�1��ImageList1�ؼ��е�ͼ��������
    ListView1.ListItems(2).SubItems(1) = "00:00:01"
    ListView1.ListItems(2).SubItems(2) = "2009-01-01"
    ListView1.ListItems(2).SubItems(3) = "(��)"
    ListView1.ListItems(2).SubItems(4) = "192.168.0.1"
    
    ListView1.ListItems.Add , , "3", , 1                                        '���ͼ��  �����Ǹ�1��ImageList1�ؼ��е�ͼ��������
    ListView1.ListItems(3).SubItems(1) = "00:00:01"
    ListView1.ListItems(3).SubItems(2) = "2010-01-01"
    ListView1.ListItems(3).SubItems(3) = "(��)"
    ListView1.ListItems(3).SubItems(4) = "192.168.0.20"
    
    '-------------------------------------------------------
    ListView1.View = lvwReport                                                  '������ʾ��ʽΪ�б�
    ListView1.AllowColumnReorder = True                                         '���н��г������У�������������
    ListView1.Arrange = lvwAutoLeft                                             'ͼ�������
    ListView1.Arrange = lvwAutoTop                                              'ͼ��������
    ListView1.FlatScrollBar = False                                             '��ʾ������
    ListView1.FlatScrollBar = True                                              '���ع�����
    ListView1.FullRowSelect = True                                              'ѡ������
    ListView1.LabelEdit = lvwManual                                             '��ֹ��ǩ�༭
    ListView1.GridLines = True                                                  '��ʾ������
    ListView1.LabelWrap = True                                                  'ͼ����Ի���
    ListView1.MultiSelect = True                                                '����ѡ������Ŀ
    ListView1.PictureAlignment = lvwTopLeft                                     'ͼƬ���뷽ʽ���󶥲����������Ҷ���(1)����ײ�(2)���ҵײ�(3)������(4)��ƽ��(5)
    ListView1.Checkboxes = False                                                '��ʾ��ѡ��
    'ListView1.DropHighlight = ListView1.ListItems.Item(2)                       '��ʾϵͳ��ɫ
    
    
    '-----------------�ѵڶ�������Ϊ��ɫ
    Dim index As Integer, line As Integer
    line = 2
    For index = 1 To ListView1.ListItems(line).ListSubItems.Count
        ListView1.ListItems(line).ListSubItems.Item(index).ForeColor = vbRed
    Next
    ListView1.ListItems(line).ForeColor = vbRed
    
    '--------------------------------------------------
    '�ѵ����� �ڶ�������Ϊ��ɫ
    ListView1.ListItems(3).ListSubItems.Item(2).ForeColor = vbBlue
    
End Sub












