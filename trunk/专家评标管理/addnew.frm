VERSION 5.00
Begin VB.Form frmAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����ר����Ϣ-����ר�ҹ���"
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
      Caption         =   "ȡ��"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "����"
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   3840
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "����ר����Ϣ"
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
            Name            =   "����"
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
         Text            =   "��"
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
         Caption         =   "ר�ҷ��ࣺ"
         Height          =   180
         Index           =   12
         Left            =   3840
         TabIndex        =   27
         Top             =   2880
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "��    ע��"
         Height          =   180
         Index           =   11
         Left            =   120
         TabIndex        =   25
         Top             =   3360
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "�ֻ����룺"
         Height          =   180
         Index           =   10
         Left            =   120
         TabIndex        =   23
         Top             =   2880
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "סլ�绰��"
         Height          =   180
         Index           =   9
         Left            =   3840
         TabIndex        =   22
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "��λ�绰��"
         Height          =   180
         Index           =   8
         Left            =   120
         TabIndex        =   21
         Top             =   2400
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "������λ��"
         Height          =   180
         Index           =   7
         Left            =   3840
         TabIndex        =   20
         Top             =   1920
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ר    ҵ��"
         Height          =   180
         Index           =   6
         Left            =   120
         TabIndex        =   19
         Top             =   1920
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ְ    �ƣ�"
         Height          =   180
         Index           =   5
         Left            =   3840
         TabIndex        =   18
         Top             =   1440
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "�Ļ��̶ȣ�"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "�������£�"
         Height          =   180
         Index           =   3
         Left            =   3840
         TabIndex        =   16
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "�Ա�"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "������"
         Height          =   180
         Index           =   1
         Left            =   3840
         TabIndex        =   5
         Top             =   480
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "��ţ�"
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
        MsgBox ("��Ų���Ϊ�գ�")
        Exit Sub
    ElseIf Trim(txtXingming.Text) = "" Then
        MsgBox ("��������Ϊ�գ�")
        Exit Sub
    ElseIf Trim(cmbXingbie.Text) = "" Then
        MsgBox ("�Ա���Ϊ�գ�")
        Exit Sub
    ElseIf Trim(txtChushengriqi.Text) = "" Then
        MsgBox ("�������²���Ϊ�գ�")
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
