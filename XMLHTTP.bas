Attribute VB_Name = "XMLHTTP"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Function GetDataEx(ByVal Url As String) As String
  On Error GoTo ERR:
  Dim XMLHTTP As Object
  Set XMLHTTP = CreateObject("Microsoft.XMLHTTP")
  XMLHTTP.open "GET", Url, True
  XMLHTTP.send
  While XMLHTTP.ReadyState <> 4
  Sleep 10
    DoEvents
  Wend
    GetDataEx = XMLHTTP.ResponseText
  Set XMLHTTP = Nothing
  Exit Function
ERR:
  GetDataEx = ""
End Function
