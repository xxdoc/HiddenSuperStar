Attribute VB_Name = "ModuleSMM2"
Option Explicit
'HiddenSuperStar By YidaozhanYa
'API: https://tgrcode.com/mm2/docs/
'定义


'变量
Public SavePath As String, ShowPicture As Boolean, CountPerPage As String
Attribute ShowPicture.VB_VarUserMemId = 1073741824
Attribute CountPerPage.VB_VarUserMemId = 1073741824
'常量
Public TGRCODE_API As String
Attribute TGRCODE_API.VB_VarUserMemId = 1073741827

'函数
Public Function GetCourseMeta(CoursePath As String) As String()
On Error GoTo errHandler
    If CheckFileExists(CoursePath) = False Then Exit Function
    '0=关卡名
    '1=简介
    '2=游戏风格
    '3=场景
    '4=版本
    '5=纪录
    '6=条件 英文
    '7=卷轴
    '8=水面
    Dim CourseMeta(0 To 8) As String, arrJSON As Variant, strMD5 As String
    With New MD5Hash
        strMD5 = LCase(.HashBytes(StrConv(CoursePath, vbFromUnicode)))
    End With
    If CheckFileExists(EnvTempDir & "\HSSTemp\" & strMD5 & ".json") = False Then
        ShellAndWait """" & App.Path & "\toost\bin\toost.exe" & """ --overworldJson """ & EnvTempDir & "\HSSTemp\" & strMD5 & "_orig.json" & """ -p """ & CoursePath & """"
        ShellAndWait """" & App.Path & "\" & IconvWrapperFilename & ".bat"" " & strMD5
        Kill EnvTempDir & "\HSSTemp\" & strMD5 & "_orig.json"
    End If
    arrJSON = Split(Left(ReadTextFile(EnvTempDir & "\HSSTemp\" & strMD5 & ".json"), 1000), ",")
    CourseMeta(0) = GetFirstObject(arrJSON, "name", CoursePath)
    CourseMeta(1) = GetFirstObject(arrJSON, "description", CoursePath)
    CourseMeta(2) = GetFirstObject(arrJSON, "gamestyle", CoursePath)
    CourseMeta(3) = ConvertLevelTheme(GetFirstObject(arrJSON, "theme", CoursePath))
    CourseMeta(4) = GetFirstObject(arrJSON, "game_version", CoursePath)
    CourseMeta(5) = CStr(GetFirstObject(arrJSON, "clear_time", CoursePath))
    If Len(CourseMeta(5)) = 6 Then
        CourseMeta(5) = CInt(Left(CourseMeta(5), 3)) \ 60 & ":" & Right(Left(CourseMeta(5), 3), 2) & "." & Right(CourseMeta(5), 3)
    ElseIf Len(CourseMeta(5)) = 5 Then
        CourseMeta(5) = "0:" & Left(CourseMeta(5), 2) & "." & Right(CourseMeta(5), 3)
    ElseIf Len(CourseMeta(5)) = 7 Then
        CourseMeta(5) = CInt(Left(CourseMeta(5), 4)) \ 60 & ":" & Right(Left(CourseMeta(5), 4), 2) & "." & Right(CourseMeta(5), 3)
    ElseIf Len(CourseMeta(5)) = 4 Then
        CourseMeta(5) = "0:0" & Left(CourseMeta(5), 1) & "." & Right(CourseMeta(5), 3)
    End If
    If CourseMeta(5) = "-1" Then CourseMeta(5) = "无"
    CourseMeta(6) = GetFirstObject(arrJSON, "clear_condition", CoursePath)
    CourseMeta(7) = GetFirstObject(arrJSON, "autoscroll_type", CoursePath) & " " & GetFirstObject(arrJSON, "autoscroll_speed", CoursePath)
    CourseMeta(8) = GetFirstObject(arrJSON, "liquid_speed", CoursePath)
    GetCourseMeta = CourseMeta
    Exit Function
errHandler:
    If Err.Number = "53" Then
        MsgBox "无法向 Temp 文件夹中写入数据，请使用 HiddenSuperStarPortable.bat 来启动本工具。"
    Else
        MsgBox "运行时错误 " & Err.Number & vbCrLf & Err.Description
    End If
    End
End Function

Public Function ConvertLevelTheme(InputStr As String) As String
    Select Case InputStr
    Case "Castle"
        ConvertLevelTheme = "城堡"
    Case "Airship"
        ConvertLevelTheme = "飞船"
    Case "Ghost house"
        ConvertLevelTheme = "鬼屋"
    Case "Underground"
        ConvertLevelTheme = "地下"
    Case "Sky"
        ConvertLevelTheme = "天空"
    Case "Snow"
        ConvertLevelTheme = "雪原"
    Case "Desert"
        ConvertLevelTheme = "沙漠"
    Case "Overworld"
        ConvertLevelTheme = "地面"
    Case "Forest"
        ConvertLevelTheme = "森林"
    Case "Underwater"
        ConvertLevelTheme = "水中"
    Case Else
        ConvertLevelTheme = InputStr
    End Select
End Function
Public Function PrettyID(ID As String) As String
    PrettyID = Left(ID, 3) & "-" & Right(Left(ID, 6), 3) & "-" & Right(ID, 3)
End Function
Private Function GetFirstObject(JSONArray As Variant, ObjToGet As String, CoursePath As String) As String
On Error GoTo errHandler
    GetFirstObject = Replace(Split(Replace(Replace(Filter(JSONArray, Chr(34) & ObjToGet & Chr(34))(0), "{", ""), "}", ""), ":")(1), Chr(34), "")
    Debug.Print GetFirstObject
    Exit Function
errHandler:
    'Kill CoursePath
    MsgBox "存档里含有已损坏或过大的关卡数据 " & CoursePath & " ，已删除。" & vbCrLf & "请重新打开本程序。", vbCritical
    End
End Function
Public Function PrettyTag(TagID As Integer) As String
    Select Case TagID
    Case 1
        PrettyTag = "标准"
    Case 2
        PrettyTag = "解谜"
    Case 3
        PrettyTag = "计时挑战"
    Case 4
        PrettyTag = "自动卷轴"
    Case 5
        PrettyTag = "自动马力欧"
    Case 6
        PrettyTag = "一次通过"
    Case 7
        PrettyTag = "多人对战"
    Case 8
        PrettyTag = "机关设计"
    Case 9
        PrettyTag = "音乐"
    Case 10
        PrettyTag = "美术"
    Case 11
        PrettyTag = "技巧"
    Case 12
        PrettyTag = "射击"
    Case 13
        PrettyTag = "BOSS战"
    Case 14
        PrettyTag = "单打"
    Case 15
        PrettyTag = "林克"
    End Select
End Function
