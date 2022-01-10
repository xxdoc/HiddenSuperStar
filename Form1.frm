VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "隐藏の无敌星"
   ClientHeight    =   1245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5040
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1245
   ScaleWidth      =   5040
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "下载"
      Height          =   615
      Left            =   3720
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtLevelPath 
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox txtLevelID 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "存档路径："
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "关卡 ID:"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LevelPath As String, Completed As Boolean
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Private Sub Command1_Click()
Command1.Caption = "下载中"
Const UrlBase = "https://tgrcode.com/mm2/level_data/"
Dim LevelID As String
LevelID = UCase(Replace(Replace(txtLevelID.Text, "-", ""), " ", ""))
If Len(LevelID) <> 9 Then
MsgBox "关卡 ID 有误。", vbOKOnly + vbExclamation, "错误"
Exit Sub
End If
DoEvents

LevelPath = txtLevelPath.Text
If Right(LevelPath, 1) = "\" Or Right(LevelPath, 1) = "/" Then
LevelPath = Left(LevelPath, Len(LevelPath) - 1)
End If
DoEvents

Open App.Path & "\config.txt" For Output As #2
Print #2, LevelPath
Close #2

Dim FileName As String, d As String
d = Dir(LevelPath & "\*.bcd")
FileName = LevelPath & "\" & d
Do Until d = ""
    If FileDateTime(LevelPath & "\" & d) > FileDateTime(FileName) Then FileName = LevelPath & "\" & d
    d = Dir
Loop
DoEvents
Dim Url As String
Dim LevelText As String
Url = UrlBase & LevelID
LevelText = GetDataEx(Url)
DoEvents
If VarType(LevelText) = vbString And InStr(1, LevelText, "err") <> 0 Then

If InStr(1, LevelText, "No course with that ID") <> 0 Then
MsgBox "关卡 ID 不存在。", vbOKOnly + vbExclamation, "错误"
Exit Sub
ElseIf InStr(1, LevelText, "a maker") <> 0 Then
MsgBox "这是一个玩家 ID。", vbOKOnly + vbExclamation, "错误"
Else
MsgBox LevelText, vbOKOnly + vbExclamation, "错误"
End If
Else
DoEvents
Kill FileName
URLDownloadToFile 0, Url, FileName, 0, 0
JSONStr = GetDataEx("https://tgrcode.com/mm2/level_info/" & LevelID)

DoEvents
Command1.Caption = "下载"
MsgBox JSONParse("name", JSONStr) & " (" & LevelID & ")" & vbCrLf & " 下载完成！", vbOKOnly + vbInformation, "提示"
End If

DoEvents

End Sub

Public Function CheckFileExists(FilePath As String) As Boolean
    On Error GoTo ERR
    If Len(FilePath) < 2 Then CheckFileExists = False: Exit Function
            If Dir$(FilePath, vbAllFileAttrib) <> vbNullString Then CheckFileExists = True
    Exit Function
ERR:
    CheckFileExists = False
End Function
Private Sub Form_Load()
Completed = False

If CheckFileExists(App.Path & "\config.txt") Then
Open App.Path & "\config.txt" For Input As #1
Line Input #1, LevelPath
Close #1
txtLevelPath.Text = LevelPath
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
Private Sub Form_Initialize()
InitCommonControls
End Sub
Public Function JSONParse(ByVal JSONPath As String, ByVal JSONString As String) As Variant
Dim JSON As Object
Set JSON = CreateObject("MSScriptControl.ScriptControl")
JSON.Language = "JScript"
JSONParse = JSON.eval("JSON=" & JSONString & ";JSON." & JSONPath & ";")
Set JSON = Nothing
End Function
