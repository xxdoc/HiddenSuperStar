VERSION 5.00
Begin VB.Form frmSettings 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设置"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5430
   BeginProperty Font 
      Name            =   "微软雅黑 Light"
      Size            =   10.5
      Charset         =   134
      Weight          =   290
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "设置"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   5430
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdClearCache 
      Caption         =   "清除缓存"
      Height          =   420
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存设置"
      Height          =   420
      Left            =   4200
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   420
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   5175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "缓存大小: "
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "API 服务器地址 (https://tgrcode.com 或其镜像网站)"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClearCache_Click()
ShellAndWait "cmd /c del /s /q ""%TEMP%\HSSTemp\*.json"""
    Label2.Caption = "缓存大小: " & GetFolderSize(Environ("Temp") & "\HSSTemp") & "MB"
End Sub

Private Sub cmdSave_Click()
TGRCODE_API = Text1.text
WriteIni "HiddenSuperStar", "ServerURL", TGRCODE_API, App.Path & "\Config.ini"
MsgBox "保存成功！"
Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "隐藏の无敌星 " & App.Major & "." & App.Minor & "." & App.Revision & " - 设置"
    Me.Show
    DoEvents
    Text1.text = TGRCODE_API
    Label2.Caption = "缓存大小: " & GetFolderSize(Environ("Temp") & "\HSSTemp") & "MB"
    DoEvents
    'MsgBox GetFolderSize(Environ("Temp") & "\HSSTemp")
End Sub

Private Function GetFolderSize(FolderPath As String) As String
Dim fso As New FileSystemObject, folder As Object
Set folder = fso.GetFolder(FolderPath)
GetFolderSize = Format((folder.Size / 1024) / 1024, "0.00")
End Function
