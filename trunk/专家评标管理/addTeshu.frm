VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTeshu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����ר�ҹ���ϵͳ - ����ר��"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9150
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command3 
      Caption         =   "�� ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   5040
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�� ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6000
      TabIndex        =   2
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�� ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      TabIndex        =   1
      Top             =   5040
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "addTeshu.frx":0000
      Height          =   4335
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   7646
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   18
      TabAction       =   1
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   14
      BeginProperty Column00 
         DataField       =   "id"
         Caption         =   "���"
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
         DataField       =   "bianhao"
         Caption         =   "���"
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
         DataField       =   "xingming"
         Caption         =   "����"
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
         DataField       =   "xingbie"
         Caption         =   "�Ա�"
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
         DataField       =   "chushengriqi"
         Caption         =   "��������"
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
         DataField       =   "xueli"
         Caption         =   "�Ļ��̶�"
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
         DataField       =   "zhicheng"
         Caption         =   "ְ��"
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
         DataField       =   "zhuanye"
         Caption         =   "רҵ"
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
         DataField       =   "danwei"
         Caption         =   "������λ"
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
         DataField       =   "danweidianhua"
         Caption         =   "��λ�绰"
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
         DataField       =   "zhuzhaidianhua"
         Caption         =   "סլ�绰"
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
         DataField       =   "shouji"
         Caption         =   "�ֻ�"
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
         DataField       =   "beizhu"
         Caption         =   "��ע"
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
      BeginProperty Column13 
         DataField       =   "leibie"
         Caption         =   "ר����"
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
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   824.882
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   569.764
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1604.976
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   2310.236
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
         BeginProperty Column11 
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1755.213
         EndProperty
         BeginProperty Column13 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "��ѡ��Ҫָ�����������ר�ң�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3570
   End
End
Attribute VB_Name = "frmTeshu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public zhidingType As String

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    
    If Me.zhidingType = "zhiding" Then
        frmMain.zhidingList.AddItem "���-" & DataGrid1.Columns("���") & "-����-" & DataGrid1.Columns("����") & "-�绰-" & DataGrid1.Columns("�ֻ�") & "-ר����-" & DataGrid1.Columns("ר����") & "(" & DataGrid1.Columns("���") & ")"
        sql = "update `user` set zhiding='1' where id=" & DataGrid1.Columns("���")
    ElseIf Me.zhidingType = "buzhiding" Then
        frmMain.buzhidingList.AddItem "���-" & DataGrid1.Columns("���") & "-����-" & DataGrid1.Columns("����") & "-�绰-" & DataGrid1.Columns("�ֻ�") & "-ר����-" & DataGrid1.Columns("ר����") & "(" & DataGrid1.Columns("���") & ")"
        sql = "update `user` set zhiding='0' where id=" & DataGrid1.Columns("���")
    End If
    
    '�������ݿ�
    
    frmMain.conn.Execute sql
    
    MsgBox "��ӳɹ�", , "��ܰ��ʾ"
    
End Sub

Private Sub Command3_Click()
        frmMain.Adodc1.Recordset.Resync
        frmMain.Adodc1.Recordset.Close
        
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
        
        frmMain.Adodc1.Recordset.Open sql
        
        Set Me.DataGrid1.DataSource = frmMain.Adodc1
End Sub

Private Sub Form_Load()
        sql = "select * from `user` "
        
        frmMain.Adodc1.ConnectionString = frmLogin.driverPath
        
        frmMain.Adodc1.CommandType = adCmdText
        
        frmMain.Adodc1.RecordSource = sql
        
        frmMain.Adodc1.Refresh
        
        
        Set DataGrid1.DataSource = frmMain.Adodc1
        
        DataGrid1.ReBind
        
        DataGrid1.Refresh
        
End Sub
