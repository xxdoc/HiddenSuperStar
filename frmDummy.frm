VERSION 5.00
Begin VB.Form frmDummy 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4665
   BeginProperty Font 
      Name            =   "微软雅黑 Light"
      Size            =   10.5
      Charset         =   134
      Weight          =   290
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDummy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   4665
   StartUpPosition =   2  '屏幕中心
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "正在下载中... 请耐心等待"
      BeginProperty Font 
         Name            =   "微软雅黑 Light"
         Size            =   12
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "frmDummy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Me.Caption = "隐藏の无敌星 " & App.Major & "." & App.Minor & "." & App.Revision & " - 加载"
End Sub

Public Sub DownloadWindow()
Label1 = "正在下载中... 请耐心等待"
Me.Show
DoEvents
End Sub
Public Sub ProcessWindow()
Label1 = "正在解析本地关卡数据... 请稍后"
Me.Show
DoEvents
End Sub
