VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "评标专家管理系统"
   ClientHeight    =   8025
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   11295
   Icon            =   "main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   11295
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame2 
      Caption         =   "抽取评标专家"
      Height          =   5775
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   8055
      Begin VB.Frame Frame6 
         Caption         =   "抽取结果："
         Height          =   3255
         Left            =   240
         TabIndex        =   20
         Top             =   2400
         Visible         =   0   'False
         Width           =   7575
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   10
            Left            =   840
            Top             =   2640
         End
         Begin VB.CommandButton Command6 
            Caption         =   "导出结果"
            Height          =   495
            Left            =   6000
            TabIndex        =   22
            Top             =   2640
            Width           =   1455
         End
         Begin VB.ListBox listJieguo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   2235
            ItemData        =   "main.frx":0ECA
            Left            =   120
            List            =   "main.frx":0ECC
            TabIndex        =   21
            Top             =   240
            Width           =   7335
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   240
            Top             =   2640
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
      Begin VB.TextBox txtGeshu 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5280
         TabIndex        =   26
         Text            =   "3"
         Top             =   1560
         Width           =   735
      End
      Begin VB.ComboBox groupList 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         ItemData        =   "main.frx":0ECE
         Left            =   1680
         List            =   "main.frx":0ED5
         TabIndex        =   23
         Top             =   1560
         Width           =   2535
      End
      Begin VB.CommandButton Command5 
         Caption         =   "开始滚动"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6240
         TabIndex        =   19
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "个数："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4320
         TabIndex        =   25
         Top             =   1560
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "专家组："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   24
         Top             =   1560
         Width           =   1440
      End
      Begin VB.Label lblGundong 
         Alignment       =   2  'Center
         Caption         =   "编号：XXXX 姓名：XXXX"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   7515
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "系统设置"
      Height          =   5775
      Left            =   360
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   8055
      Begin VB.ListBox zhidingList 
         Height          =   1500
         Left            =   2640
         TabIndex        =   35
         Top             =   360
         Width           =   3975
      End
      Begin VB.CommandButton Command10 
         Caption         =   "删除"
         Height          =   495
         Left            =   6720
         TabIndex        =   34
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton Command9 
         Caption         =   "增加"
         Height          =   495
         Left            =   6720
         TabIndex        =   33
         Top             =   2160
         Width           =   1095
      End
      Begin VB.CommandButton Command8 
         Caption         =   "删除"
         Height          =   495
         Left            =   6720
         TabIndex        =   32
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Command7 
         Caption         =   "增加"
         Height          =   495
         Left            =   6720
         TabIndex        =   31
         Top             =   360
         Width           =   1095
      End
      Begin VB.ListBox buzhidingList 
         Height          =   1500
         ItemData        =   "main.frx":0EDF
         Left            =   2640
         List            =   "main.frx":0EE6
         TabIndex        =   30
         Top             =   2160
         Width           =   3975
      End
      Begin VB.TextBox txtDengluMima 
         Height          =   375
         Left            =   2760
         TabIndex        =   28
         Top             =   4080
         Width           =   1935
      End
      Begin VB.TextBox txtShezhiMima 
         Height          =   375
         Left            =   2760
         TabIndex        =   15
         Top             =   4680
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "保存"
         Height          =   495
         Left            =   2760
         TabIndex        =   13
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "指定不抽取专家："
         Height          =   300
         Index           =   0
         Left            =   900
         TabIndex        =   29
         Top             =   2040
         Width           =   1440
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "登录密码："
         Height          =   180
         Left            =   1320
         TabIndex        =   27
         Top             =   4080
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "指定抽取专家："
         Height          =   180
         Index           =   2
         Left            =   1080
         TabIndex        =   16
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "系统设置密码："
         Height          =   180
         Index           =   1
         Left            =   960
         TabIndex        =   14
         Top             =   4680
         Width           =   1260
      End
      Begin VB.Line Line1 
         X1              =   2520
         X2              =   2520
         Y1              =   240
         Y2              =   5640
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "专家管理"
      Height          =   5775
      Left            =   1200
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   8055
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   450
         Left            =   6720
         Top             =   5160
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   794
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "main.frx":0EF9
         Height          =   4815
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   8493
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   2
         RowHeight       =   18
         TabAction       =   1
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "bianhao"
            Caption         =   "编号"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "xingming"
            Caption         =   "姓名"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "xingbie"
            Caption         =   "性别"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "chushengriqi"
            Caption         =   "出生年月"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "xueli"
            Caption         =   "文化程度"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "zhicheng"
            Caption         =   "职称"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "zhuanye"
            Caption         =   "专业"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "danwei"
            Caption         =   "工作单位"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "danweidianhua"
            Caption         =   "单位电话"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "zhuzhaidianhua"
            Caption         =   "住宅电话"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "shouji"
            Caption         =   "手机"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "beizhu"
            Caption         =   "备注"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column12 
            DataField       =   "leibie"
            Caption         =   "类别"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   450.142
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1604.976
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   2310.236
            EndProperty
            BeginProperty Column08 
            EndProperty
            BeginProperty Column09 
            EndProperty
            BeginProperty Column10 
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1755.213
            EndProperty
            BeginProperty Column12 
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command4 
         Caption         =   "新增"
         Height          =   495
         Left            =   4080
         TabIndex        =   17
         Top             =   5160
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   5160
         Width           =   2895
      End
      Begin VB.CommandButton Command2 
         Caption         =   "查找"
         Height          =   495
         Left            =   3120
         TabIndex        =   10
         Top             =   5160
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   10815
      Begin VB.Label Label2 
         Caption         =   "专家评标管理软件"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   36
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2280
         TabIndex        =   4
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5775
      Left            =   8520
      TabIndex        =   1
      Top             =   1680
      Width           =   2535
      Begin VB.CommandButton Command1 
         Caption         =   "系统设置"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "专家管理"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "专家评标"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Label Label1 
      Caption         =   "技术支持： QQ 23463790 E-Mail:23463790@qq.com"
      Height          =   375
      Left            =   9120
      TabIndex        =   2
      Top             =   7560
      Width           =   2055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long

Private islogin As Boolean

Private infoList() As String

Private cishu, chongfu2 As Integer

Public conn As ADODB.Connection   '建立一个新的连接对象。
Public myset As ADODB.Recordset     '建立一个新的记录集对象。

Private Sub randomcheck()
    '随机抽取专家
    Frame2.Caption = "抽取评标专家"
    Frame2.Visible = True
    Frame4.Visible = False
    Frame5.Visible = False
    Frame2.Left = 240
    Frame2.Top = 1680
    
    Call Form_Load
    
End Sub

Private Sub manager()
    '专家管理

    Frame5.Caption = "专家管理"
    Frame5.Visible = True
    Frame2.Visible = False
    Frame4.Visible = False
    Frame5.Left = 240
    Frame5.Top = 1680
    
    Text1.Text = ""
    
    
    '初始加载数据
    Call Command2_Click
    
    
End Sub

Private Sub systemset()
    '系统设置
    
    If islogin Or checkpassword() Then

        Frame4.Caption = "系统设置"
        Frame4.Visible = True
        Frame2.Visible = False
        Frame5.Visible = False
        Frame4.Left = 240
        Frame4.Top = 1680
        
    Else
    
        Exit Sub
            
    End If
    
    Call loadSystemSet
    
End Sub

Private Sub loadSystemSet()

    If txtGeshu.Text = "" Then
        txtGeshu.Text = "3"
    End If
        
    '读取系统表
    sql = "select top 1 * from `system` order by id desc"
    
    myset.Open sql, conn
        
    If Not myset.EOF Then
        
        myset.MoveFirst
        
        txtDengluMima.Text = myset("denglumima")
        txtShezhiMima.Text = myset("shezhimima")
        
    Else
    
        myset.AddNew
        myset("denglumima") = ""
        myset("shezhimima") = ""
        myset.Update
    
    End If
    
    myset.Close
    
    '读取指定的专家信息
    zhidingList.Clear
    
    buzhidingList.Clear
    
     sql = "select * from `user` where zhiding=1 or zhiding=0 order by id desc"
     
     myset.Open sql, conn
     
     Do While Not myset.EOF
            
        If myset("zhiding") = 1 Then
        
            zhidingList.AddItem "编号-" & myset("bianhao") & "-姓名-" & myset("xingming") & "-电话-" & myset("shouji") & "-专家组-" & myset("leibie") & "(" & myset("id") & ")"
            
        ElseIf myset("zhiding") = 0 Then
            
            buzhidingList.AddItem "编号-" & myset("bianhao") & "-姓名-" & myset("xingming") & "-电话-" & myset("shouji") & "-专家组-" & myset("leibie") & "(" & myset("id") & ")"
            
        End If
        
        myset.MoveNext
        
    Loop
    
    myset.Close
    
    Adodc1.ConnectionString = frmLogin.driverPath
        
    Adodc1.CommandType = adCmdText
    
    sql = "select * from `user` "
    
    Adodc1.RecordSource = sql
    
    Adodc1.Refresh
    
End Sub

Private Sub Command1_Click(index As Integer)

    If index = 0 Then
    
        Call randomcheck
    
    ElseIf index = 1 Then
    
        Call manager
        
    ElseIf index = 2 Then
    
        Call systemset
        
    End If
    
    
End Sub


Private Sub Command10_Click()
    If buzhidingList.ListIndex > -1 Then
        id = Mid(buzhidingList.Text, InStr(buzhidingList.Text, "("))
        sql = "update `user` set zhiding=null where id in " & id
        Me.conn.Execute sql
        
        buzhidingList.RemoveItem (buzhidingList.ListIndex)
        
        MsgBox "删除成功", , "温馨提示"
        
    End If
End Sub

Private Sub Command2_Click()
    
    
    If Trim(Text1.Text) <> "" Then
    
        Adodc1.Recordset.Resync
        Adodc1.Recordset.Close
        
        sql = "select * from `user` where bianhao like '%" + Text1.Text + "%'"
        
        sql = sql + " or xingming like '%" + Text1.Text + "%'"
        
        sql = sql + " or xingbie like '%" + Text1.Text + "%'"
        
        sql = sql + " or chushengriqi like '%" + Text1.Text + "%'"
        
        sql = sql + " or xueli like '%" + Text1.Text + "%'"
        
        sql = sql + " or zhicheng like '%" + Text1.Text + "%'"
        
        sql = sql + " or zhuanye like '%" + Text1.Text + "%'"
        
        sql = sql + " or danwei like '%" + Text1.Text + "%'"
        
        sql = sql + " or danweidianhua like '%" + Text1.Text + "%'"
        
        sql = sql + " or zhuzhaidianhua like '%" + Text1.Text + "%'"
        
        sql = sql + " or shouji like '%" + Text1.Text + "%'"
        
        sql = sql + " or beizhu like '%" + Text1.Text + "%'"
        
        sql = sql + " or leibie like '%" + Text1.Text + "%'"
        
        Adodc1.Recordset.Open sql
        
        Set DataGrid1.DataSource = Adodc1
    
    Else
    
        If Not Adodc1.Recordset Is Nothing Then
            Adodc1.Recordset.Resync
            Adodc1.Recordset.Close
        End If
        
        sql = "select * from `user` "
        
        Adodc1.ConnectionString = frmLogin.driverPath
        
        Adodc1.CommandType = adCmdText
        
        Adodc1.RecordSource = sql
        
        Adodc1.Refresh
        
        
        Set DataGrid1.DataSource = Adodc1
        
        DataGrid1.ReBind
        
        DataGrid1.Refresh
        

    End If
    
End Sub

Private Sub Command3_Click()
    
    sql = "select top 1 * from `system` order by id desc"
    myset.Open sql, conn, 1, 3
    
    myset.MoveFirst
    
    If myset.EOF Then
        
        myset.AddNew
        myset("denglumima") = ""
        myset("shezhimima") = ""
        myset.Update
        
    Else
    
        myset("denglumima") = txtDengluMima.Text
        myset("shezhimima") = txtShezhiMima.Text
        myset.Update
        
        MsgBox ("保存成功！")
        
    End If
    
    myset.Close
    
End Sub


Private Function checkpassword()

    '读取系统表
    sql = "select top 1 * from `system` order by id desc"
    
    myset.Open sql
    
    myset.MoveFirst
    
    If Not myset.EOF Then
    
        mima = myset("shezhimima")
    
        If islogin = True Then
        
            checkpassword = True
            
            islogin = True
            
        Else
            If mima = InputBox("请输入密码：", "登录") Then
        
                checkpassword = True
            
                islogin = True
            Else
            
                MsgBox ("密码错误！")
                checkpassword = False
                islogin = False
            End If
        
        End If
        
    End If
    
    myset.Close
    
End Function


Private Sub Command4_Click()
    frmAdd.Show
End Sub


Private Sub start(group)
    '读取专家表
    sql = "select * from `user` "
    
    If Trim(group) <> "" And Trim(group) <> "全部" And Trim(group) <> "所有" Then
        sql = sql + " where leibie like '%" & group & "%'"
    End If
    
    myset.Open sql, conn, 3, 2
    
    If Not myset.EOF Then
    
        ReDim infoList(0 To myset.RecordCount - 1)
        
        myset.MoveFirst
        
        For i = 0 To myset.RecordCount - 1 Step 1
            
            infoList(i) = "编号-" & myset("bianhao") & "-姓名-" & myset("xingming") & "-电话-" & myset("shouji") & "-专家组-" & myset("leibie") & "(" & myset("id") & ")"
            
            myset.MoveNext
            
            If myset.EOF Then Exit For
            
        Next i
        
        Timer1.Enabled = True
        
    End If
    
    myset.Close
    
End Sub

Private Sub Command5_Click()

    Erase infoList  '清空数组
    
    If (IsNumeric(txtGeshu.Text) = False) Then
        MsgBox "个数必须为大于 0 的整数！", , "友情提示"
        Exit Sub
    End If
    
    If (txtGeshu.Text < 1) Then
        
        MsgBox "个数必须为大于 0 的整数！", , "友情提示"
        Exit Sub
    
    End If
    
    Call start(groupList.Text) '读取数据并滚动字幕
    
    If SafeArrayGetDim(infoList) = 0 Then
        MsgBox "没有读取到要抽取的专家信息，请检查专家组或个数是否正确。", , "友情提示"
        Exit Sub
    End If
    
    members = UBound(infoList) - LBound(infoList) + 1 '读取到的专家数
    
    If Me.Command5.Caption = "开始滚动" Then
    
        Me.txtGeshu.Locked = True
        Me.groupList.Locked = True
        
        Me.Command5.Caption = "随机抽取"
        
        Me.listJieguo.Clear
    
    Else
        
        Me.Frame6.Visible = True
        
        '指定的专家
        If zhidingList.ListCount > 0 Then
        
            For i = 0 To zhidingList.ListCount - 1
            
                If Trim(groupList.Text) <> "" And Trim(groupList.Text) <> "全部" And Trim(groupList.Text) <> "所有" Then
                    If InStr(zhidingList.List(i), groupList.Text) > 0 Then
                        listJieguo.AddItem zhidingList.List(i)
                    End If
                Else
                    listJieguo.AddItem zhidingList.List(i)
                End If
                                
                If listJieguo.ListCount = Me.txtGeshu.Text Then Exit For
                                
            Next i
            
        End If
        
        
        
        For i = 0 To Int(txtGeshu.Text) - 1
            
            If members > 0 Then
            
                If listJieguo.ListCount >= Int(txtGeshu.Text) Then Exit For
                
                n = Int(Rnd * members)
                
                isBuzhiding = True
                
                For k = 0 To buzhidingList.ListCount - 1
                    
                    If InStr(buzhidingList.List(k), infoList(n)) > 0 Then
                    
                        infoList = removeArray(infoList, n)
                
                        members = members - 1
                    
                        isBuzhiding = False
                        
                        Exit For
                    
                    End If
                    
                Next k
                
                For k = 0 To listJieguo.ListCount - 1
                    
                    If listJieguo.List(k) = infoList(n) Then
                        
                        infoList = removeArray(infoList, n)
                
                        members = members - 1
                        
                        isBuzhiding = False
                        
                        Exit For
                        
                    End If
                    
                Next k
                
                If isBuzhiding Then
                    
                    listJieguo.AddItem infoList(n)
                    
                    infoList = removeArray(infoList, n)
                
                    members = members - 1
                    
                Else
                
                    i = i - 1
                
                End If
                            
            End If
                            
        Next i
        
        Me.txtGeshu.Locked = False
        Me.groupList.Locked = False
        
        Me.Timer1.Enabled = False
        Me.Command5.Caption = "开始滚动"
        
        lblGundong.Alignment = 2
        lblGundong.Caption = "编号：XXXX 姓名：XXXX"
        
        
    End If
    
    
End Sub

Private Sub Command6_Click()
    Dim filename_select, apppath As String
    
    apppath = App.Path
    
    CommonDialog1.DialogTitle = "请选择要保存的位置："
    CommonDialog1.InitDir = "c:\" '缺省打开路径
    CommonDialog1.Filter = "文本文件*.txt|*.txt" '过滤器
    CommonDialog1.FileName = "专家评标抽取结果.txt"
    CommonDialog1.ShowSave 'showopen这个是打开，如果要保存的话改成commondialog1.showsave好了
    
    
    If CommonDialog1.FileName <> "" Then
        Open CommonDialog1.FileName For Output As #1
        For i = 0 To listJieguo.ListCount - 1
            Print #1, Left(listJieguo.List(i), InStr(listJieguo.List(i), "(") - 1)
        Next i
        Close #1
    End If
    
    ChDir apppath
    
End Sub

Private Sub Command7_Click()
    frmTeshu.zhidingType = "zhiding"
    frmTeshu.Show
    
End Sub

Private Sub Command8_Click()
    
    If zhidingList.ListIndex > -1 Then
        id = Mid(zhidingList.Text, InStr(zhidingList.Text, "("))
        sql = "update `user` set zhiding=null where id in " & id
        Me.conn.Execute sql
        
        zhidingList.RemoveItem (zhidingList.ListIndex)
        
        MsgBox "删除成功", , "温馨提示"
        
    End If
    
End Sub

Private Sub Command9_Click()
    frmTeshu.zhidingType = "buzhiding"
    frmTeshu.Show
End Sub

Private Sub DataGrid1_BeforeDelete(Cancel As Integer)
    r = MsgBox("确定要删除记录吗？", vbYesNo, "警告")
    If r = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub Form_Load()

    
    Set conn = New ADODB.Connection   '建立一个新的连接对象。
    Set myset = New ADODB.Recordset      '建立一个新的记录集对象。
    
    conn.Open frmLogin.driverPath
    
    
    '加载系统参数
    Call loadSystemSet
    
    
     '读取专家组信息
    
    
    sql = "select leibie from `user` group by leibie order by leibie asc"
    
    myset.Open sql, conn
    
    If Not myset.EOF Then
    
        groupList.Clear
    
        myset.MoveFirst
        
        Do While Not myset.EOF
            
            If myset("leibie") <> "" Then
            
                groupList.AddItem (myset("leibie"))
                
            End If
            
            myset.MoveNext
            
        Loop
    
    End If
    
    myset.Close
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub listJieguo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    listJieguo.ToolTipText = listJieguo.Text
End Sub

Private Sub Timer1_Timer()
    
    Dim m As Integer
    
    m = UBound(infoList) - LBound(infoList) + 1

    If m >= 0 Then
        
         n = Int(Rnd * m)
         
         lblGundong.Caption = infoList(n)
         
         lblGundong.Alignment = n Mod 3
         
    Else
    
        Timer1.Enabled = False
        
    End If
    
End Sub


Private Function removeArray(arr As Variant, ByRef index As Variant)
    counts = UBound(arr) - LBound(arr)
    Dim tempArr() As String
    ReDim tempArr(counts)
    
    For i = 0 To counts
        If i = counts Then Exit For
        If i >= index Then
            tempArr(i) = arr(i + 1)
        Else
            tempArr(i) = arr(i)
        End If
    Next
    removeArray = tempArr
End Function

