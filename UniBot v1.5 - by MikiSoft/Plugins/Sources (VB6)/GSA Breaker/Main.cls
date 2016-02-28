VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GSABreaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IPluginInterface
Dim Program As Form, WithEvents Requests As cAsyncRequests, CaptchaText As New Collection
Attribute Requests.VB_VarHelpID = -1

Private Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal _
        dest As Long, ByVal src As Long, ByVal Length As Long) As Long

Private Sub CatBinary(bytData() As Byte, Bytes() As Byte)
    Dim BytesLen As Long, BinaryNext As Long
    
    BinaryNext = UBound(bytData) + 1
    BytesLen = UBound(Bytes) - LBound(Bytes) + 1
    If BinaryNext + BytesLen > BinaryNext Then ReDim Preserve bytData(BinaryNext + BytesLen - 1)
    CopyMemory VarPtr(bytData(BinaryNext)), VarPtr(Bytes(LBound(Bytes))), BytesLen
End Sub

Private Sub CatBinaryString(bytData() As Byte, Text As String)
    Dim Bytes() As Byte

    Bytes = StrConv(Text, vbFromUnicode)
    CatBinary bytData, Bytes
End Sub

Private Sub RandStr(strT As String)
    Dim strR As String: strR = "0123456789qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM"
    Dim i As Byte
    For i = 1 To 15
        strT = strT & Mid$(strR, Int(Rnd() * Len(strR) + 1), 1)
    Next
End Sub

Private Sub IPluginInterface_BuildSettings(strS As String)
End Sub

Private Function IPluginInterface_Startup(Parent As Object) As String
Set Program = Parent
IPluginInterface_Startup = "gsa"
Set Requests = New cAsyncRequests
End Function

Private Sub Class_Terminate()
Set Requests = Nothing
End Sub

Private Function IPluginInterface_Info(Inf() As String) As Object
Inf(0) = App.Major & "." & App.Minor
Inf(1) = "MikiSoft"
Inf(2) = "Combines UniBot with GSA Captcha Breaker local server (http://127.0.0.1/gsa_test.gsa)." & vbNewLine & _
"Example: gsa('http://site.com/captchaimage','useragent','cookiedata','referer') - where parameters after URL are optional"
End Function

Private Function IPluginInterface_Execute(Params() As String) As String
Dim CurrKey As String: CurrKey = "-" & Requests.RequestCount
On Error GoTo Error
CaptchaText.Add vbNullChar, CurrKey
Dim strH As String
If UBound(Params) >= 2 Then
If Params(2) <> vbNullString Then strH = "User-Agent" & vbLf & Params(2)
If UBound(Params) >= 3 Then
If Params(3) <> vbNullString Then strH = strH & vbLf & "Cookie" & vbLf & Params(3)
If UBound(Params) = 4 Then If Params(4) <> vbNullString Then strH = strH & vbLf & "Referer" & vbLf & Params(4)
End If
End If
Requests.AddRequest(CurrKey).SendRequest Params(1), , strH
Do While GetItem(CurrKey) = vbNullChar
DoEvents
Loop
CurrKey = Mid$(CurrKey, 2) & "-"
Dim CapTxt As String: CapTxt = GetItem(CurrKey)
If CapTxt <> vbNullString Then
IPluginInterface_Execute = CapTxt
Requests.RemoveRequest CurrKey
CaptchaText.Remove CurrKey
Exit Function
End If
Error:
On Error GoTo -1
On Error Resume Next
If Requests.RequestCount > 0 Then Requests.RemoveRequest CurrKey
Program.addLog "(GSA Breaker) [" & Replace(CurrKey, "-", vbNullString, , 1) + 1 & "] Error!"
End Function

Private Function GetItem(Key As String) As String
On Error Resume Next
GetItem = CaptchaText.Item(Key)
End Function

Private Sub Requests_ResponseFinished(Req As cAsyncRequest)
Dim TmpStr As String
On Error GoTo E
If StrPtr(Req.Tag) = 0 Then
RandStr TmpStr
Dim PostData() As Byte: PostData = ""
CatBinaryString PostData, "--" & TmpStr & vbCrLf & "Content-Disposition: form-data; name=""file""; filename=""captcha.png""" & vbCrLf & "Content-Type: " & Req.http.GetResponseHeader("Content-Type") & vbCrLf & vbCrLf
CatBinary PostData, Req.http.ResponseBody
CatBinaryString PostData, vbCrLf & "--" & TmpStr & vbCrLf & "Content-Disposition: form-data; name=""action""" & vbCrLf & vbCrLf & "Submit" & vbCrLf & "--" & TmpStr & "--" & vbCrLf
Req.Tag = ""
Program.addLog "(GSA Breaker) [" & Mid$(Left$(Req.Key, InStr(Req.Key & ",", ",") - 1), 2) + 1 & "] Potential captcha downloaded. Solving captcha..."
Req.SendRequest "http://127.0.0.1/gsa_test.gsa", PostData, "Content-Type" & vbLf & "multipart/form-data; boundary=" & TmpStr
Erase PostData
Exit Sub
Else
TmpStr = Req.http.ResponseText
Dim Res As Integer: Res = InStr(TmpStr, "<span id=""captcha_result"">")
If Res = 0 Then GoTo E
TmpStr = Mid$(TmpStr, Res + 26)
TmpStr = Left$(TmpStr, InStr(TmpStr, "<") - 1)
TmpStr = Replace(Replace(Replace(TmpStr, "&gt;", ">"), "&lt;", "<"), "&amp;", "&")
If TmpStr <> vbNullString Then CaptchaText.Add TmpStr, Mid$(Req.Key, 2) & "-"
End If
E: CaptchaText.Remove Req.Key
End Sub

Private Sub Requests_Error(Req As cAsyncRequest, ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
CaptchaText.Remove Req.Key
End Sub
