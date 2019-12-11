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
   StartUpPosition =   3  '窗口缺省
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

'单击 ImageList1 控件 右键 -> 属性 进行添加图标

Private Sub Form_Load()
    
    ListView1.SmallIcons = ImageList1.Object                                    '把ImageList1 控件绑定到 ListView1列表控件中来。 下面需要用ImageList1中的图标
    
    ListView1.ListItems.Clear                                                   '清空列表
    ListView1.ColumnHeaders.Clear                                               '清空列表头
    ListView1.View = lvwReport                                                  '设置列表显示方式
    ListView1.GridLines = True                                                  '显示网络线
    ListView1.LabelEdit = lvwManual                                             '禁止标签编辑
    ListView1.FullRowSelect = True                                              '选择整行
    
    ListView1.ColumnHeaders.Add , , "ID", 500                                   '给列表中添加列名
    ListView1.ColumnHeaders.Add , , "本地 IP", 1500
    ListView1.ColumnHeaders.Add , , "本地端口", 1200
    ListView1.ColumnHeaders.Add , , "协议", 550
    ListView1.ColumnHeaders.Add , , "远程 IP", 1500
    ListView1.ColumnHeaders.Add , , "远程端口", 900
    ListView1.ColumnHeaders.Add , , "当前状态", 900
    ListView1.ColumnHeaders.Add , , "连接时间", 900
    '-------------------------------------------------------
    Dim X
    X = ListView1.ListItems.Count + 1
    ListView1.ListItems.Add , , X
    ListView1.ListItems(X).SubItems(1) = "00:00:00"
    ListView1.ListItems(X).SubItems(2) = "2008-01-01"
    ListView1.ListItems(X).SubItems(3) = "(无)"
    '-------------------------------------------------------
    ListView1.ListItems.Clear                                                   '清空列表
    ListView1.ListItems.Add , , "1", , 1                                        '添加图标  后面那个1是ImageList1控件中的图标索引号
    ListView1.ListItems(1).SubItems(1) = "00:00:00"
    ListView1.ListItems(1).SubItems(2) = "2008-01-01"
    ListView1.ListItems(1).SubItems(3) = "(无)"
    ListView1.ListItems(1).SubItems(4) = "127.0.0.1"
    
    ListView1.ListItems.Add , , "2", , 2                                        '添加图标  后面那个1是ImageList1控件中的图标索引号
    ListView1.ListItems(2).SubItems(1) = "00:00:01"
    ListView1.ListItems(2).SubItems(2) = "2009-01-01"
    ListView1.ListItems(2).SubItems(3) = "(无)"
    ListView1.ListItems(2).SubItems(4) = "192.168.0.1"
    
    ListView1.ListItems.Add , , "3", , 1                                        '添加图标  后面那个1是ImageList1控件中的图标索引号
    ListView1.ListItems(3).SubItems(1) = "00:00:01"
    ListView1.ListItems(3).SubItems(2) = "2010-01-01"
    ListView1.ListItems(3).SubItems(3) = "(无)"
    ListView1.ListItems(3).SubItems(4) = "192.168.0.20"
    
    '-------------------------------------------------------
    ListView1.View = lvwReport                                                  '设置显示方式为列表
    ListView1.AllowColumnReorder = True                                         '对行进行程序排列，用鼠标进行排列
    ListView1.Arrange = lvwAutoLeft                                             '图标横排列
    ListView1.Arrange = lvwAutoTop                                              '图标竖排列
    ListView1.FlatScrollBar = False                                             '显示滚动条
    ListView1.FlatScrollBar = True                                              '隐藏滚动条
    ListView1.FullRowSelect = True                                              '选择整行
    ListView1.LabelEdit = lvwManual                                             '禁止标签编辑
    ListView1.GridLines = True                                                  '显示网络线
    ListView1.LabelWrap = True                                                  '图标可以换行
    ListView1.MultiSelect = True                                                '可以选择多个项目
    ListView1.PictureAlignment = lvwTopLeft                                     '图片对齐方式是左顶部，其他有右顶部(1)、左底部(2)、右底部(3)、居中(4)、平铺(5)
    ListView1.Checkboxes = False                                                '显示复选框
    'ListView1.DropHighlight = ListView1.ListItems.Item(2)                       '显示系统颜色
    
    
    '-----------------把第二行设置为红色
    Dim index As Integer, line As Integer
    line = 2
    For index = 1 To ListView1.ListItems(line).ListSubItems.Count
        ListView1.ListItems(line).ListSubItems.Item(index).ForeColor = vbRed
    Next
    ListView1.ListItems(line).ForeColor = vbRed
    
    '--------------------------------------------------
    '把第三行 第二列设置为蓝色
    ListView1.ListItems(3).ListSubItems.Item(2).ForeColor = vbBlue
    
End Sub












