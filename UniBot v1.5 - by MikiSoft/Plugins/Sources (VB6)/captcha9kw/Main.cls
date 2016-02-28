VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_NAME = "captcha9kw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal _
dest As Long, ByVal src As Long, ByVal Length As Long) As Long

Implements IPluginInterface
Dim Program As Form, WithEvents Requests As cAsyncRequests, CaptchaText As New Collection, bolH As Boolean, Prms(9) As String
Attribute Requests.VB_VarHelpID = -1

Const SERVER As String = "www.9kw.eu"
Const SOFT_KEY As String = "unibot"

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

Private Function GetItem(Key As String) As String
On Error Resume Next
GetItem = CaptchaText.Item(Key)
End Function

Private Sub Class_Terminate()
If Program Is Nothing Then Exit Sub
Set Requests = Nothing
If Not frmS.bolSave Then Exit Sub
Dim strT As String: IPluginInterface_BuildSettings strT
On Error Resume Next
If strT <> vbNullString Then
Open App.Path & "\" & App.EXEName & ".ini" For Output Access Write As #1
Print #1, strT;
Close #1
If bolH Then SetAttr App.Path & "\" & App.EXEName & ".ini", vbHidden
Else: Kill App.Path & "\" & App.EXEName & ".ini"
End If
End Sub

Private Sub IPluginInterface_BuildSettings(strS As String)
If Not frmS.bolSave Then Exit Sub
If frmS.strKey = vbNullString And frmS.bytMin = 3 And frmS.bytMaxL = 0 And frmS.bytCase = 0 And frmS.bytPrior = 0 And Not frmS.bolMath And Not frmS.bolPhrase And Not frmS.bolNum And Not frmS.bolOCR And Not frmS.bolNoSpace And Not frmS.bolNoRem Then Exit Sub
If frmS.strKey <> vbNullString Then strS = "key=" & frmS.strKey & vbCrLf
If frmS.bytMin <> 3 Then strS = strS & "minl=" & frmS.bytMin & vbCrLf
If frmS.bytMaxL > 0 Then strS = strS & "maxl=" & frmS.bytMaxL & vbCrLf
If frmS.bytCase > 0 Then strS = strS & "case=" & frmS.bytCase & vbCrLf
If frmS.bytPrior > 0 Then strS = strS & "priority=1" & vbCrLf
If frmS.bolMath Then strS = strS & "math=1" & vbCrLf
If frmS.bolPhrase Then strS = strS & "phrase=1" & vbCrLf
If frmS.bolNum Then strS = strS & "numeric=1" & vbCrLf
If frmS.bolOCR Then strS = strS & "ocr=1" & vbCrLf
If frmS.bolNoSpace Then strS = strS & "nospace=1" & vbCrLf
If frmS.bolNoRem Then strS = strS & "nochng=1" & vbCrLf
strS = strS & "save=1" & vbCrLf
strS = Left$(strS, Len(strS) - 2)
End Sub

Private Function IPluginInterface_Startup(Parent As Object) As String
Set Program = Parent
IPluginInterface_Startup = "9kw"
Set Requests = New cAsyncRequests
frmS.bytMin = 3
Prms(0) = "min_len"
Prms(1) = "max_len"
Prms(2) = "case-sensitive"
Prms(3) = "phrase"
Prms(4) = "numeric"
Prms(5) = "math"
Prms(6) = "ocr"
Prms(7) = "nospace"
Prms(8) = "prio"
Prms(9) = "maxtimeout"
On Error GoTo E
If Dir$(App.Path & "\" & App.EXEName & ".ini", vbHidden) <> vbNullString Then
bolH = GetAttr(App.Path & "\" & App.EXEName & ".ini") = vbHidden
Dim strL(1) As String
Open App.Path & "\" & App.EXEName & ".ini" For Input Access Read As #1
While Not EOF(1)
Line Input #1, strL(0)
strL(0) = Trim$(Replace(strL(0), vbCr, vbNullString))
If strL(0) <> vbNullString Then
If InStr(";#[", Left$(strL(0), 1)) = 0 Then
strL(1) = Left$(strL(0), InStr(strL(0), "=") - 1)
strL(0) = Mid$(Left$(strL(0), Len(strL(0))), Len(strL(1)) + 2)
If IsNumeric(strL(0)) Then
If strL(0) < 0 Then strL(0) = strL(0) * (-1)
Select Case strL(1)
Case "minl": If strL(0) < 255 Then frmS.bytMin = strL(0) Else: frmS.bytMin = 255
Case "maxl": If strL(0) < 255 Then frmS.bytMaxL = strL(0) Else: frmS.bytMaxL = 255
Case "case": If strL(0) < 2 Then frmS.bytCase = strL(0) Else: frmS.bytCase = 2
Case "priority": If strL(0) < 20 Then frmS.bytPrior = strL(0) Else: frmS.bytPrior = 20
Case "math": frmS.bolMath = strL(0) > 0
Case "phrase": frmS.bolPhrase = strL(0) > 0
Case "numeric": frmS.bolNum = strL(0) > 0
Case "nospace": frmS.bolNoSpace = strL(0) > 0
Case "ocr": frmS.bolOCR = strL(0) > 0
Case "nochng": frmS.bolNoRem = strL(0) > 0
Case "save": frmS.bolSave = strL(0) > 0
End Select
ElseIf strL(1) = "key" Then frmS.strKey = frmS.RplKey(strL(0))
End If
End If
End If
Wend
Close #1
ChngM frmS.bytMin, frmS.bytMaxL, 3
End If
E:
'On Error GoTo -1
'On Error Resume Next
'If frmS.strKey <> vbNullString Then Exit Function
'If Right$(Program.lblStatus.Caption, 7) = "Idle..." Then
'MsgBox "You haven't set a valid API key yet!", vbExclamation
'frmS.Show vbModal, Program
'ElseIf Program.Visible Then MsgBox "You haven't set a valid API key yet! Go to settings of this plugin and do that before using it.", vbExclamation
'End If
End Function

Private Sub ChngM(bytM As Byte, bytM1 As Byte, bytDMin As Byte)
If bytM = 0 Then
If bytM1 > 0 Then bytM = bytM1 Else: bytM = bytDMin
ElseIf bytM1 > 0 And bytM1 < bytM Then bytM1 = bytM
End If
End Sub

Private Function IPluginInterface_Info(Inf() As String) As Object
Inf(0) = App.Major & "." & App.Minor
Inf(1) = "MikiSoft"
Inf(2) = "Combines UniBot with 9kw.eu service." & vbNewLine & _
"Example: 9kw('http://site.com/captchaimage','useragent','cookiedata','referer','apikey',3,0,0,0,0,0,0,0,0,5,60)" & vbNewLine & _
"Parameters after API key: Min. & max. length, Case sensitive (0 - not specified, 1, 2 - with symbols), Phrase (0-1), Numeric (0-1), Math (0-1), OCR (0-1), No space (0-1), Priority, Response wait, Max. response time." & vbNewLine & _
"All parameters after URL are optional."
Set IPluginInterface_Info = frmS
End Function

Private Function IPluginInterface_Execute(Params() As String) As String
Dim CurrKey As String, strK As String
CurrKey = "-" & Requests.RequestCount
On Error GoTo Error
Dim strH As String, i As Byte
If UBound(Params) >= 2 Then
If Not IsNumeric(Params(2)) Then
If Params(2) <> vbNullString Then strH = "User-Agent" & vbLf & Params(2)
If UBound(Params) >= 3 Then
If Not IsNumeric(Params(3)) Then
If Params(3) <> vbNullString Then strH = strH & vbLf & "Cookie" & vbLf & Params(3)
If UBound(Params) >= 4 Then
If Not IsNumeric(Params(4)) Then
If Params(4) <> vbNullString Then strH = strH & vbLf & "Referer" & vbLf & Params(4)
If UBound(Params) >= 5 Then
If Not IsNumeric(Params(5)) Then
strK = frmS.RplKey(Params(5))
If strK <> vbNullString Then CurrKey = CurrKey & "," & strK Else: GoSub C1
If UBound(Params) >= 6 Then i = 6
Else: i = 5
End If
End If
Else: i = 4
End If
End If
Else: i = 3
End If
End If
Else: i = 2
End If
If UBound(Params) < 5 Or InStr(CurrKey, ",") = 0 Then GoSub C1
If i > 0 Then
Dim bolT As Boolean
GoSub InpM
If Params(i) = "0" Or Params(i) = "1" Or Params(i) = "2" Then CurrKey = CurrKey & "," & Params(i) Else: CurrKey = CurrKey & ","
For i = i + 1 To i + 5
If UBound(Params) < i Then GoTo C
If Params(i) = "0" Or Params(i) = "1" Then CurrKey = CurrKey & "," & Params(i) Else: CurrKey = CurrKey & ","
Next
If UBound(Params) >= i Then
If IsNumeric(Params(i)) Then
If Params(i) < 0 Then Params(i) = Params(i) * (-1)
If Params(i) > 20 Then Params(i) = 20
CurrKey = CurrKey & "," & Params(i)
Else: CurrKey = CurrKey & ","
End If
i = i + 1
GoSub InpM
End If
End If
Else
C1:
If frmS.strKey = vbNullString Then
Program.addLog "(9kw) There is no API key set!"
Exit Function
ElseIf i > 0 Then
CurrKey = CurrKey & "," & frmS.strKey
Return
End If
End If
C:
CaptchaText.Add vbNullChar, CurrKey
Requests.AddRequest(CurrKey).SendRequest Params(1), , strH
Do While GetItem(CurrKey) = vbNullChar
DoEvents
Loop
Requests.RemoveRequest CurrKey
CurrKey = Mid$(CurrKey, 2) & "-"
Dim CapTxt As String: CapTxt = GetItem(CurrKey)
If CapTxt <> vbNullString Then
IPluginInterface_Execute = CapTxt
CaptchaText.Remove CurrKey
End If
Exit Function
Error:
On Error GoTo -1
On Error Resume Next
If Requests.RequestCount > 0 Then Requests.RemoveRequest CurrKey
Program.addLog "(9kw) [" & Replace(Left$(CurrKey, InStr(CurrKey & ",", ",") - 1), "-", vbNullString, , 1) + 1 & "] Error!"
Exit Function
InpM:
bolT = False
For i = i To i + 1
If UBound(Params) < i Then GoTo C
If IsNumeric(Params(i)) Then
If Params(i) < 0 Then Params(i) = Params(i) * (-1)
If Params(i) > 255 Then Params(i) = 255
If Not bolT Then bolT = True Else: If CByte(Params(i)) < CByte(Params(i - 1)) Then Params(i) = Params(i - 1)
CurrKey = CurrKey & "," & Params(i)
ElseIf i >= 13 Then CurrKey = CurrKey & ","
End If
Next
Return
End Function

Private Sub Requests_ResponseFinished(Req As cAsyncRequest)
Dim TmpStr(1) As String, bytT(1) As Byte, i As Byte
On Error GoTo E
TmpStr(1) = Req.Key
If StrPtr(Req.Tag) = 0 Then
RandStr TmpStr(0)
Dim PostData() As Byte
PostData = ""
bytT(0) = InStr(TmpStr(1), ",") + 1
If bytT(0) > 1 Then
For i = 0 To 9
bytT(1) = InStr(bytT(0), TmpStr(1), ",")
If bytT(1) = 0 Then bytT(1) = Len(TmpStr(1)) + 1
If bytT(1) > bytT(0) Then
If i > 0 Then
bytT(0) = Mid$(TmpStr(1), bytT(0), bytT(1) - bytT(0))
If Not frmS.bolNoRem Then
Select Case i
Case 1: frmS.txtMin = bytT(0)
Case 2: frmS.txtMaxL = bytT(0)
Case 3: frmS.bytCase = bytT(0)
Case 4: frmS.bolPhrase = bytT(0) = 1
Case 5: frmS.bolNum = bytT(0) = 1
Case 6: frmS.bolMath = bytT(0) = 1
Case 7: frmS.bolOCR = bytT(0) = 1
Case 8: frmS.bolNoSpace = bytT(0) = 1
Case 9: frmS.bytPrior = bytT(0)
End Select
End If
If bytT(0) > 0 Then CatBinaryString PostData, "--" & TmpStr(0) & vbCrLf & "Content-Disposition: form-data; name=""" & Prms(i - 1) & """" & vbCrLf & vbCrLf & bytT(0) & vbCrLf
Else
Dim strK As String: strK = Mid$(TmpStr(1), bytT(0), bytT(1) - bytT(0))
If Not frmS.bolNoRem Then frmS.strKey = strK
End If
End If
If bytT(1) < Len(TmpStr(1)) Then bytT(0) = bytT(1) + 1 Else: Exit For
Next
End If
Dim CT As String: CT = Req.http.GetResponseHeader("Content-Type")
CatBinaryString PostData, "--" & TmpStr(0) & vbCrLf & "Content-Disposition: form-data; name=""source""" & vbCrLf & vbCrLf & SOFT_KEY & vbCrLf & _
"--" & TmpStr(0) & vbCrLf & "Content-Disposition: form-data; name=""username""" & vbCrLf & vbCrLf & strK & vbCrLf & _
"--" & TmpStr(0) & vbCrLf & "Content-Disposition: form-data; name=""action""" & vbCrLf & vbCrLf & "UPLOADCAPTCHA" & vbCrLf & _
"--" & TmpStr(0) & vbCrLf & "Content-Disposition: form-data; name=""file""; filename=""captcha." & Mid$(CT, InStr(CT, "/") + 1) & """" & vbCrLf & "Content-Type: " & CT & vbCrLf & vbCrLf
CatBinary PostData, Req.http.ResponseBody
CatBinaryString PostData, vbCrLf & "--" & TmpStr(0) & "--" & vbCrLf
Program.addLog "(9kw) [" & Mid$(Left$(TmpStr(1), InStr(TmpStr(1) & ",", ",") - 1), 2) + 1 & "] Potential captcha downloaded. Submitting captcha..."
Req.Tag = ""
Req.SendRequest "http://" & SERVER & "/index.cgi", PostData, "Content-Type" & vbLf & "multipart/form-data; boundary=" & TmpStr(0)
Erase PostData
Else
TmpStr(0) = Req.http.ResponseText
Debug.Print TmpStr(0)
If Left$(TmpStr(0), 2) <> "00" Then
Dim Res As Integer: Res = InStr(TmpStr(0), "<span id=""captcha_result"">")
If Res = 0 Then GoTo E
TmpStr(0) = Mid$(TmpStr(0), Res + 26)
TmpStr(0) = Left$(TmpStr(0), InStr(TmpStr(0), "<") - 1)
TmpStr(0) = Replace(Replace(Replace(TmpStr(0), "&gt;", ">"), "&lt;", "<"), "&amp;", "&")
If TmpStr(0) <> vbNullString Then CaptchaText.Add TmpStr(0), Mid$(Req.Key, 2) & "-"
CaptchaText.Remove Req.Key
Else
Select Case Mid$(TmpStr(0), 3, 2)
Case "01", "02", "03", "04": TmpStr(0) = "The key doesn't exist"
Case "08", "09", "13": TmpStr(0) = "Captcha couldn't be solved"
Case "11", "24": TmpStr(0) = "No credit on the account"
Case "10": TmpStr(0) = "Image size isn't allowed"
Case "31": TmpStr(0) = "Account isn't yet 24 hours present in the system"
Case "32": TmpStr(0) = "Account doesn't have the full rights"
Case "15": TmpStr(0) = "Captcha submitted too quickly"
End Select
TmpStr(0) = vbNullChar & TmpStr(0)
GoTo E
End If
End If
Exit Sub
E:
If Left$(TmpStr(0), 1) <> vbNullChar Then TmpStr(0) = "!" Else: TmpStr(0) = ": " & Mid$(TmpStr(0), 2) & "."
Program.addLog "(9kw) [" & Mid$(Left$(TmpStr(1), InStr(TmpStr(1) & ",", ",") - 1), 2) + 1 & "] Error" & TmpStr(0)
CaptchaText.Remove Req.Key
End Sub

Private Sub Requests_Error(Req As cAsyncRequest, ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
Program.addLog "(9kw) [" & Mid$(Left$(Req.Key, InStr(Req.Key & ",", ",") - 1), 2) + 1 & "] Error!"
CaptchaText.Remove Req.Key
End Sub
