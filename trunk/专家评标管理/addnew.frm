VERSION 5.00
Begin VB.Form frmAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "新增专家信息-评标专家管理"
   ClientHeight    =   4800
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7470
   Icon            =   "addnew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "保存"
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   3840
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "新增专家信息"
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7215
      Begin VB.TextBox txtBeizhu 
         Height          =   855
         Left            =   1080
         TabIndex        =   28
         Top             =   3360
         Width           =   2175
      End
      Begin VB.TextBox txtLeibie 
         Height          =   375
         Left            =   4800
         TabIndex        =   26
         Top             =   2760
         Width           =   2055
      End
      Begin VB.TextBox txtShouji 
         Height          =   375
         Left            =   1080
         TabIndex        =   24
         Top             =   2760
         Width           =   2175
      End
      Begin VB.ComboBox cmbXingbie 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "addnew.frx":0ECA
         Left            =   1080
         List            =   "addnew.frx":0ED4
         TabIndex        =   14
         Text            =   "男"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtZhuzhaidianhua 
         Height          =   375
         Left            =   4800
         TabIndex        =   13
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox txtDanweidianhua 
         Height          =   375
         Left            =   1080
         TabIndex        =   12
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox txtDanwei 
         Height          =   375
         Left            =   4800
         TabIndex        =   11
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox txtZhuanye 
         Height          =   375
         Left            =   1080
         TabIndex        =   10
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox txtZhicheng 
         Height          =   375
         Left            =   4800
         TabIndex        =   9
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txtXueli 
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox txtChushengriqi 
         Height          =   375
         Left            =   4800
         TabIndex        =   7
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtXingming 
         Height          =   375
         Left            =   4800
         TabIndex        =   6
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtBianhao 
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "专家分类："
         Height          =   180
         Index           =   12
         Left            =   3840
         TabIndex        =   27
         Top             =   2880
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "备    注："
         Height          =   180
         Index           =   11
         Left            =   120
         TabIndex        =   25
         Top             =   3360
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "手机号码："
         Height          =   180
         Index           =   10
         Left            =   120
         TabIndex        =   23
         Top             =   2880
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "住宅电话："
         Height          =   180
         Index           =   9
         Left            =   3840
         TabIndex        =   22
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "单位电话："
         Height          =   180
         Index           =   8
         Left            =   120
         TabIndex        =   21
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "工作单位："
         Height          =   180
         Index           =   7
         Left            =   3840
         TabIndex        =   20
         Top             =   1920
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "专    业："
         Height          =   180
         Index           =   6
         Left            =   120
         TabIndex        =   19
         Top             =   1920
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "职    称："
         Height          =   180
         Index           =   5
         Left            =   3840
         TabIndex        =   18
         Top             =   1440
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "文化程度："
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "出生年月："
         Height          =   180
         Index           =   3
         Left            =   3840
         TabIndex        =   16
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "性别："
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "姓名："
         Height          =   180
         Index           =   1
         Left            =   3840
         TabIndex        =   5
         Top             =   480
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "编号："
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub



Private Sub OKButton_Click()

    If Trim(txtBianhao.Text) = "" Then
        MsgBox ("编号不能为空！")
        Exit Sub
    ElseIf Trim(txtXingming.Text) = "" Then
        MsgBox ("姓名不能为空！")
        Exit Sub
    ElseIf Trim(cmbXingbie.Text) = "" Then
        MsgBox ("性别不能为空！")
        Exit Sub
    ElseIf Trim(txtChushengriqi.Text) = "" Then
        MsgBox ("出生年月不能为空！")
        Exit Sub
    End If
    
    
    frmMain.Adodc1.Recordset.AddNew
    
    frmMain.Adodc1.Recordset("bianhao") = txtBianhao.Text
    frmMain.Adodc1.Recordset("xingming") = txtXingming.Text
    frmMain.Adodc1.Recordset("xingbie") = cmbXingbie.Text
    frmMain.Adodc1.Recordset("chushengriqi") = txtChushengriqi.Text
    frmMain.Adodc1.Recordset("xueli") = txtXueli.Text
    frmMain.Adodc1.Recordset("zhicheng") = txtZhicheng.Text
    frmMain.Adodc1.Recordset("zhuanye") = txtZhuanye.Text
    frmMain.Adodc1.Recordset("danwei") = txtDanwei.Text
    frmMain.Adodc1.Recordset("danweidianhua") = txtDanweidianhua.Text
    frmMain.Adodc1.Recordset("zhuzhaidianhua") = txtZhuzhaidianhua.Text
    frmMain.Adodc1.Recordset("shouji") = txtShouji.Text
    frmMain.Adodc1.Recordset("leibie") = txtLeibie.Text
    frmMain.Adodc1.Recordset("beizhu") = txtBeizhu.Text
    
    frmMain.Adodc1.Recordset.Update
    
    Unload Me
End Sub
