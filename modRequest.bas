Attribute VB_Name = "modRequest"
Option Explicit

Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Private Declare Function OleLoadPicturePath Lib "oleaut32.dll" (ByVal szURLorPath As Long, ByVal punkCaller As Long, ByVal dwReserved As Long, ByVal clrReserved As OLE_COLOR, ByRef riid As TGUID, ByRef ppvRet As IPicture) As Long
Private Type TGUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type


Public Function xhrGet(url As String) As String
    Dim xhr As Object
    Set xhr = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    xhr.Open "GET", url, True
    xhr.send
    xhr.waitForResponse 60
    If xhr.Status = 200 Then
        xhrGet = xhr.responseText
    ElseIf xhr.Status = 400 Then
        MsgBox "你输入的ID类型错误。", vbCritical
        xhrGet = "Error"
    ElseIf xhr.Status = 404 Then
        MsgBox "未找到这个关卡或玩家。", vbCritical
        xhrGet = "Error"
    Else
        MsgBox "HTTP错误 " & xhr.Status & " " & xhr.statusText, vbCritical
        xhrGet = ""
        End
    End If
    Set xhr = Nothing
End Function

Public Function LoadPicture(ByVal strFileName As String) As Picture
    Dim IID As TGUID
    With IID
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    On Error GoTo LocalErr
    OleLoadPicturePath StrPtr(strFileName), 0&, 0&, 0&, IID, LoadPicture
    Exit Function
LocalErr:
    Set LoadPicture = VB.LoadPicture(strFileName)
    Err.Clear
End Function
