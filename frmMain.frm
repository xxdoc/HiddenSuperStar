VERSION 5.00
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.2#0"; "VBCCR17.OCX"
Object = "{A2A736C2-8DAC-4CDB-B1CB-3B077FBB14F9}#6.2#0"; "VB6Resizer2.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H80000005&
   Caption         =   "隐藏的无敌星 - 本地关卡"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10320
   BeginProperty Font 
      Name            =   "Microsoft YaHei UI Light"
      Size            =   10.5
      Charset         =   134
      Weight          =   290
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   10320
   StartUpPosition =   3  '窗口缺省
   Begin VB6ResizerLib2.VB6Resizer VB6Resizer1 
      Left            =   7080
      Top             =   4560
      _ExtentX        =   529
      _ExtentY        =   529
   End
   Begin VB.CommandButton cmdSettings 
      Caption         =   "设置"
      Height          =   495
      Left            =   9120
      TabIndex        =   6
      Tag             =   "TL"
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdShowPicture 
      Caption         =   "查看关卡图片"
      Height          =   495
      Left            =   5760
      TabIndex        =   5
      Tag             =   "TL"
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "浏览在线关卡"
      Default         =   -1  'True
      Height          =   495
      Left            =   5760
      TabIndex        =   4
      Tag             =   "TL"
      Top             =   6480
      Width           =   1575
   End
   Begin VBCCR17.ListView lstLocal 
      Height          =   6855
      Left            =   120
      TabIndex        =   3
      Tag             =   "HW"
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   12091
      SmallIcons      =   "imgLstIcons"
      View            =   3
   End
   Begin VBCCR17.ImageList imgLstIcons 
      Left            =   4440
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      ImageWidth      =   16
      ImageHeight     =   16
      UseBackColor    =   -1  'True
      InitListImages  =   "frmMain.frx":54AA
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "关于"
      Height          =   495
      Left            =   9120
      TabIndex        =   2
      Tag             =   "TL"
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdChangeSave 
      Caption         =   "更改存档路径"
      Height          =   495
      Left            =   7440
      TabIndex        =   0
      Tag             =   "TL"
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Label lblAboutCourse 
      BackStyle       =   0  'Transparent
      Caption         =   "请选择一个关卡"
      Height          =   5535
      Left            =   5760
      TabIndex        =   1
      Tag             =   "HL"
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LoadCompletedM As Boolean
Private Sub cmdAbout_Click()
    frmAbout.Show
End Sub

Private Sub cmdBrowse_Click()
    frmBrowse.Show
End Sub

Private Sub cmdChangeSave_Click()
    SelectSavePath
End Sub


Private Sub cmdSettings_Click()
frmSettings.Show
End Sub

Private Sub cmdShowPicture_Click()
    ShellAndWait Chr(34) & App.Path & "\toost\bin\toost.exe" & Chr(34) & " -o " & Chr(34) & Environ("Temp") & "\HSSTemp\表世界.png" & Chr(34) & " -s " & Chr(34) & Environ("Temp") & "\HSSTemp\里世界.png" & Chr(34) & " -p " & Chr(34) & SavePath & "\" & lstLocal.SelectedItem.Tag & Chr(34)
    Shell "cmd /c start " & Chr(34) & Chr(34) & " " & Chr(34) & Environ("Temp") & "\HSSTemp\表世界.png" & Chr(34)
    Shell "cmd /c start " & Chr(34) & Chr(34) & " " & Chr(34) & Environ("Temp") & "\HSSTemp\里世界.png" & Chr(34)
End Sub

Private Sub Form_Activate()
On Error Resume Next
cmdBrowse.SetFocus
End Sub

Private Sub Form_Initialize()
    InitCommonControls
    '加载
    'TGRCODE_API = "https://tgrcode.com"
    TGRCODE_API = GetIni("HiddenSuperStar", "ServerURL", App.Path & "\Config.ini")
    'ShellAndWait "cmd /c rd /s /q " & Chr(34) & Environ("Temp") & "\HSSTemp" & Chr(34)
    If Dir(Environ("Temp") & "\HSSTemp", vbDirectory) = "" Then MkDir Environ("Temp") & "\HSSTemp"
    '检查完整性
    Dim CheckIntegrity As Boolean
    CheckIntegrity = False
    If CheckFileExists(App.Path & "\toost\bin\toost.exe") = False Then CheckIntegrity = True
    If CheckFileExists(App.Path & "\iconv.exe") = False Then CheckIntegrity = True
    If CheckFileExists(App.Path & "\iconv_wrapper.bat") = False Then CheckIntegrity = True
    If CheckFileExists(App.Path & "\Config.ini") = False And CheckFileExists(App.Path & "\Config.Defaults.ini") = False Then CheckIntegrity = True
    If CheckIntegrity Then MsgBox "检查完整性失败，请重新下载程序！", vbCritical: End
    If CheckFileExists(App.Path & "\Config.ini") = False And CheckFileExists(App.Path & "\Config.Defaults.ini") = True Then Name App.Path & "\Config.Defaults.ini" As App.Path & "\Config.ini"

End Sub

Private Sub SelectSavePath()
    Dim SavePathTmp As String
    If SavePath = "None" Then
        Do Until SavePathTmp <> ""
            SavePathTmp = ChooseDir("选择存档目录", Me)
        Loop
    Else
        SavePathTmp = ChooseDir("选择存档目录", Me)
        If SavePathTmp = "" Then Exit Sub
    End If
    If CheckFileExists(SavePathTmp & "\save.dat") = False Then GoTo Err
    SavePath = SavePathTmp
    WriteIni "HiddenSuperStar", "SavePath", SavePath, App.Path & "\Config.ini"
    Exit Sub
Err:
    MsgBox "请选择正确的存档目录。"
    If SavePath = "None" Then End
End Sub

Private Sub Form_Load()
    Me.Caption = "隐藏の无敌星 " & App.Major & "." & App.Minor & "." & App.Revision & " - 本地关卡"
        SavePath = GetIni("HiddenSuperStar", "SavePath", App.Path & "\Config.ini")
    If SavePath = "None" Then
        MsgBox "欢迎使用 隐藏の无敌星！" & vbCrLf & "在首次使用之前，你需要先选择你的《马造2》存档文件夹。" & vbCrLf & vbCrLf & _
               "你可以在模拟器中右键马造2，点击“打开存档目录”来获得存档文件夹的路径。"
        SelectSavePath
    End If
    If Dir(SavePath, vbDirectory) = "" Then
        MsgBox "欢迎使用 隐藏の无敌星！" & vbCrLf & "在首次使用之前，你需要先选择你的《马造2》存档文件夹。" & vbCrLf & vbCrLf & _
               "你可以在模拟器中右键马造2，点击“打开存档目录”来获得存档文件夹的路径。"
        SavePath = "None"
        SelectSavePath
    End If
    LoadLocalLevels
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub lstLocal_Click()
    lblAboutCourse.Caption = "关卡文件：" & lstLocal.SelectedItem.Tag
    Dim CourseMetadata() As String
    CourseMetadata = GetCourseMeta(SavePath & "\" & lstLocal.SelectedItem.Tag)
    LblAppend lblAboutCourse, CourseMetadata(0)
    LblAppend lblAboutCourse, "游戏风格: " & CourseMetadata(2)
    LblAppend lblAboutCourse, "场景：" & CourseMetadata(3)
    LblAppend lblAboutCourse, "水面: " & Replace(CourseMetadata(8), "None", "无")
    LblAppend lblAboutCourse, "卷轴: " & Replace(CourseMetadata(7), "None", "无")
    LblAppend lblAboutCourse, "纪录: " & CStr(CourseMetadata(5))
    LblAppend lblAboutCourse, "游戏版本: " & CourseMetadata(4)
    LblAppend lblAboutCourse, "过关条件: " & Replace(CourseMetadata(6), "None", "无")
    LblAppend lblAboutCourse, ""
    LblAppend lblAboutCourse, CourseMetadata(1)
End Sub


Private Sub txtCourseID_Change()
    txtCourseID.text = UCase(txtCourseID.text)
    txtCourseID.SelStart = Len(txtCourseID.text)

End Sub

'txtCourseID.ForeColor = RGB(100, 100, 100)
Private Sub txtCourseID_Click()
    If txtCourseID.text = "关卡ID" Then txtCourseID.ForeColor = RGB(0, 0, 0): txtCourseID.text = ""
End Sub

Private Sub txtCourseID_KeyPress(KeyAscii As Integer)
    If Len(txtCourseID.text) = 11 And KeyAscii <> 8 Then KeyAscii = 0
    If Len(txtCourseID.text) = 3 Or Len(txtCourseID.text) = 7 Then
        If KeyAscii <> 8 Then txtCourseID.text = txtCourseID.text & "-"
    End If
End Sub

Public Sub LoadLocalLevels()
    Dim CourseList() As String, CourseMetadata() As String
    LoadCompletedM = False
    frmDummy.ProcessWindow
    DoEvents
    lstLocal.ListItems.Clear
    lstLocal.ColumnHeaders.Clear
    lstLocal.FullRowSelect = True
    lstLocal.GridLines = True
    lstLocal.ColumnHeaders.Add 1, "Icon", "", 400
    lstLocal.ColumnHeaders.Add 2, "GameStyle", "版本", 1000
    lstLocal.ColumnHeaders.Add 3, "CourseName", "关卡", 3500
    CourseList = GetFileList(SavePath, "course_data_0*.bcd")
    Dim CC As Integer, Course As Variant
    For Each Course In CourseList
        CC = lstLocal.ListItems.Count + 1
        CourseMetadata = GetCourseMeta(SavePath & "\" & Course)
        DoEvents
        lstLocal.ListItems.Add , , "", , CourseMetadata(2)
        lstLocal.ListItems(CC).SubItems(1) = CourseMetadata(2)
        lstLocal.ListItems(CC).SubItems(2) = CourseMetadata(0)
        With lstLocal.ListItems(CC)
            .ToolTipText = CourseMetadata(0)
            .Tag = Course
        End With
    Next Course
    lstLocal.ListItems(1).Selected = False
    frmDummy.Hide
    Unload frmDummy
    LoadCompletedM = True
End Sub

Private Sub VB6Resizer1_AfterResize()
    If LoadCompletedM Then lstLocal.ColumnHeaders(3).Width = 3500 + (frmMain.Width - 10440)
End Sub

