Attribute VB_Name = "Module1"
Public Declare Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As Long, Optional ByVal pszStrPtr As Long) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()


Public Function CheckFileExists(FilePath As String) As Boolean
'检查文件是否存在
    On Error GoTo Err
    If Len(FilePath) < 2 Then CheckFileExists = False: Exit Function
    If Dir$(FilePath, vbAllFileAttrib) <> vbNullString Then CheckFileExists = True
    Exit Function
Err:
    CheckFileExists = False
End Function

Public Sub ShellAndWait(pathFile As String)
    With CreateObject("WScript.Shell")
        .Run pathFile, 0, True
    End With
End Sub

Public Function GetIni(strSection As String, strKey As String, INIFileName As String)
    With New ClassINI
        .INIFileName = INIFileName
        GetIni = .GetIniKey(strSection, strKey)
    End With
End Function

Public Sub WriteIni(strSection As String, strKey As String, strNewValue As String, INIFileName As String)
    With New ClassINI
        .INIFileName = INIFileName
        .WriteIniKey strSection, strKey, strNewValue
    End With
End Sub

Public Function ChooseDir(ByVal frmTitle As String, onForm As Object) As String
'oleexp 选择目录
    On Error Resume Next
    Dim pChooseDir As New FileOpenDialog
    Dim psiResult As IShellItem
    Dim lpPath As Long, sPath As String
    With pChooseDir
        .SetOptions FOS_PICKFOLDERS
        .SetTitle frmTitle
        .Show onForm.hWnd
        .GetResult psiResult
        If (psiResult Is Nothing) = False Then
            psiResult.GetDisplayName SIGDN_FILESYSPATH, lpPath
            If lpPath Then
                SysReAllocString VarPtr(sPath), lpPath
                CoTaskMemFree lpPath
            End If
        End If
    End With
    If BStrFromLPWStr(lpPath) <> "" Then ChooseDir = BStrFromLPWStr(lpPath)
End Function

Public Function BStrFromLPWStr(lpWStr As Long) As String
    SysReAllocString VarPtr(BStrFromLPWStr), lpWStr
End Function
Function GetFileList(ByVal Path As String, Optional fExp As String = "*.*") As String()
    Dim fName As String, i As Long, FileName() As String
    If Right$(Path, 1) <> "\" Then Path = Path & "\"
    fName = Dir$(Path & fExp)
    i = 0
    Do While fName <> ""
        ReDim Preserve FileName(i) As String
        FileName(i) = fName
        fName = Dir$
        i = i + 1
    Loop
    If i <> 0 Then
        ReDim Preserve FileName(i - 1) As String
    End If
    GetFileList = FileName
End Function

Public Function ReadTextFile(sFilePath As String) As String
    On Error Resume Next

    Dim handle As Integer
    If LenB(Dir$(sFilePath)) > 0 Then

        handle = FreeFile
        Open sFilePath For Binary As #handle
        ReadTextFile = Space$(LOF(handle))
        Get #handle, , ReadTextFile
        Close #handle

    End If

End Function

Public Sub LblAppend(label As Object, text As String)
    label.Caption = label.Caption & vbCrLf & text
End Sub
