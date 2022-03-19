VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "关于本程序"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5610
   BeginProperty Font 
      Name            =   "Microsoft YaHei UI Light"
      Size            =   10.5
      Charset         =   0
      Weight          =   290
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5610
   StartUpPosition =   3  '窗口缺省
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   3735
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   120
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Image1.Picture = Me.Icon
    Label1.Caption = "关于 隐藏の无敌星 " & App.Major & "." & App.Minor & "." & App.Revision
    LblAppend Label1, ""
    LblAppend Label1, "TheGreatRambler 马造 2 API 的图形化前端，"
    LblAppend Label1, "为 NS 模拟器专门适配。"
    LblAppend Label1, "为提升加载速度可以使用特殊网络环境，"
    LblAppend Label1, "* 此版本为测试版，非稳定版！"
    LblAppend Label1, ""
    LblAppend Label1, "2022 是一刀斩哒"
    LblAppend Label1, "API By TheGreatRambler"
End Sub

