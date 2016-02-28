VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_NAME = "p_2captcha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const SOFT_KEY As String = "912"
Private Const SERVER As String = "2captcha.com"

Implements IPluginInterface
Dim Program As Form, WithEvents Requests As cAsyncRequests, WithEvents Tmr As SelfTimer, CaptchaText As New Collection, CaptchaID As New Collection, bolH As Boolean, Prms(4) As String
Attribute Requests.VB_VarHelpID = -1
Attribute Tmr.VB_VarHelpID = -1

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

Private Function GetItem(Key As String) As String
On Error Resume Next
GetItem = CaptchaText.Item(Key)
End Function

Private Sub Class_Terminate()
If Program Is Nothing Then Exit Sub
Set Requests = Nothing
Set Tmr = Nothing
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
If frmS.strKey = vbNullString And frmS.bytType = 0 And frmS.bytLang = 0 And frmS.bytWait = 5 And frmS.bytMax = 60 And Not frmS.bolCase And Not frmS.bolMath And Not frmS.bolPhrase And Not frmS.bolNoRem And frmS.bytAfter = 0 Then Exit Sub
If frmS.strKey <> vbNullString Then strS = "key=" & frmS.strKey & vbCrLf
If frmS.bytType > 0 Then strS = strS & "type=" & frmS.bytType & vbCrLf
If frmS.bytLang > 0 Then strS = strS & "lang=" & frmS.bytLang & vbCrLf
If frmS.bytWait <> 5 Then strS = strS & "wait=" & frmS.bytWait & vbCrLf
If frmS.bytMax <> 60 Then strS = strS & "max=" & frmS.bytMax & vbCrLf
If frmS.bolCase Then strS = strS & "case=1" & vbCrLf
If frmS.bolMath Then strS = strS & "math=1" & vbCrLf
If frmS.bolPhrase Then strS = strS & "phrase=1" & vbCrLf
If frmS.bytAfter > 0 Then strS = strS & "after=" & frmS.bytAfter & vbCrLf
If frmS.bolNoRem Then strS = strS & "nochng=1" & vbCrLf
strS = strS & "save=1" & vbCrLf
strS = Left$(strS, Len(strS) - 2)
End Sub

Private Function IPluginInterface_Startup(Parent As Object) As String
Set Program = Parent
IPluginInterface_Startup = "2cap"
Set Requests = New cAsyncRequests
Set Tmr = New SelfTimer
frmS.bytWait = 5
frmS.bytMax = 60
Prms(0) = "regsense"
Prms(1) = "phrase"
Prms(2) = "calc"
Prms(3) = "numeric"
Prms(4) = "language"
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
Case "type": If strL(0) <= 3 Then frmS.bytType = strL(0)
Case "lang": If strL(0) <= 2 Then frmS.bytLang = strL(0)
Case "wait": If strL(0) < 255 Then frmS.bytWait = strL(0) Else: frmS.bytWait = 255
Case "max": If strL(0) < 255 Then frmS.bytMax = strL(0) Else: frmS.bytMax = 255
Case "case": frmS.bolCase = strL(0) > 0
Case "math": frmS.bolMath = strL(0) > 0
Case "phrase": frmS.bolPhrase = strL(0) > 0
Case "after": If strL(0) < 255 Then frmS.bytAfter = strL(0) Else: frmS.bytAfter = 255
Case "nochng": frmS.bolNoRem = strL(0) > 0
Case "save": frmS.bolSave = strL(0) > 0
End Select
ElseIf strL(1) = "key" Then frmS.strKey = frmS.RplKey(strL(0))
End If
End If
End If
Wend
Close #1
If frmS.bytWait = 0 Then
If frmS.bytMax > 0 Then frmS.bytWait = frmS.bytMax Else: frmS.bytWait = 5
Else: If frmS.bytMax > 0 And frmS.bytMax < frmS.bytWait Then frmS.bytMax = frmS.bytWait
End If
End If
E: 'If frmS.strKey = vbNullString And Right$(Program.lblStatus.Caption, 7) = "Idle..." And Program.Visible Then MsgBox "You haven't set a valid API key yet! Go to settings of this plugin and do that before using it.", vbExclamation
End Function

Private Function IPluginInterface_Info(Inf() As String) As Object
Inf(0) = App.Major & "." & App.Minor
Inf(1) = "MikiSoft"
Inf(2) = "Combines UniBot with 2captcha.com service." & vbNewLine & _
"Example: 2cap('http://site.com/captchaimage','useragent','cookiedata','referer','apikey',0,0,0,0,0,0,5,60)" & vbNewLine & _
"Parameters after API key: Store captcha for report (0-1), Case sensitive (0-1), Phrase (0-1), Math (0-1), Type (0 - not specified, 1 - numeric, 2 - letters, 3 - both), Language (0 or 1 - Latin, 2 - Cyrillic), Response wait, Max. response time." & vbNewLine & _
"All parameters after URL are optional." & vbNewLine & _
"For reporting incorrect captcha text: 2cap('captchatext'[,'apikey'])"
Set IPluginInterface_Info = frmS
End Function

Private Function IPluginInterface_Execute(Params() As String) As String
Dim CurrKey As String, strK As String, bolT As Boolean
If Left$(Params(1), 1) = "`" Or (Left$(Params(1), 7) <> "http://" And Left$(Params(1), 8) <> "https://") Then
If Params(2) <> vbNullString Then
strK = frmS.RplKey(Params(2))
If strK = vbNullString Then bolT = True: GoSub C1
Else: bolT = True: GoSub C1
End If
If Left$(Params(1), 1) = "`" Then Params(1) = Mid$(Params(1), 2)
If Tmr.Interval > 0 Then
On Error Resume Next
CurrKey = CaptchaID.Item(Params(1))
On Error GoTo 0
If CurrKey <> vbNullString Then
CaptchaID.Remove Params(1)
If CaptchaID.Count = 0 Then Tmr.Interval = 0
Requests.AddRequest(Requests.RequestCount & "-" & Params(1) & " ").SendRequest "http://" & SERVER & "/res.php?key=" & strK & "&action=reportbad&id=" & CurrKey
Else: Program.addLog "(2captcha) Unknown captcha text."
End If
Else: Program.addLog "(2captcha) Storing option isn't enabled."
End If
Exit Function
End If
CurrKey = "-" & Requests.RequestCount
On Error GoTo Error
Dim strH As String, i As Byte, intI As Integer
intI = frmS.bytAfter + 1
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
If IsNumeric(Params(i)) Then
If Params(i) < 0 Then Params(i) = Params(i) * (-1)
If Params(i) > 255 Then Params(i) = 255
intI = Params(i)
If intI > 1 And Not frmS.bolNoRem Then frmS.bytAfter = intI - 1
End If
i = i + 1
If UBound(Params) >= i Then
For i = i To i + 2
If UBound(Params) < i Then GoTo C
If Params(i) = "0" Or Params(i) = "1" Then CurrKey = CurrKey & "," & Params(i) Else: CurrKey = CurrKey & ","
Next
If UBound(Params) >= i Then
If Params(i) = "0" Or Params(i) = "1" Or Params(i) = "2" Or Params(i) = "3" Then CurrKey = CurrKey & "," & Params(i) Else: CurrKey = CurrKey & ","
i = i + 1
If UBound(Params) >= i Then
If Params(i) = "0" Or Params(i) = "1" Or Params(i) = "2" Then CurrKey = CurrKey & "," & Params(i) Else: CurrKey = CurrKey & ","
For i = i + 1 To i + 2
If UBound(Params) >= i Then
If IsNumeric(Params(i)) Then
If Params(i) < 0 Then Params(i) = Params(i) * (-1)
If Params(i) > 255 Then Params(i) = 255
If Not bolT Then bolT = True Else: If CByte(Params(i)) < CByte(Params(i - 1)) Then Params(i) = Params(i - 1)
CurrKey = CurrKey & "," & Params(i)
End If
Else: Exit For
End If
Next
End If
End If
End If
End If
Else
C1:
If frmS.strKey = vbNullString Then
Program.addLog "(2captcha) There is no API key set!"
Exit Function
ElseIf Not bolT Then
If i > 0 Then
CurrKey = CurrKey & "," & frmS.strKey
Return
End If
Else
strK = frmS.strKey
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
strH = GetItem(CurrKey)
Dim CapTxt As String
If strH <> vbNullString Then
i = InStr(strH, "|")
CapTxt = Left$(strH, i - 1)
If intI > 0 Then
On Error Resume Next
CaptchaID.Remove CapTxt
On Error GoTo Error
CaptchaID.Add Mid$(strH, i + 1), CapTxt
If intI = 1 Then If frmS.bytAfter > 0 Then intI = frmS.bytAfter + 1 Else: intI = 2
If intI > 1 Then If Tmr.Interval <> (intI - 1) * 1000 Then Tmr.Interval = (intI - 1) * 1000
End If
IPluginInterface_Execute = CapTxt
CaptchaText.Remove CurrKey
End If
Exit Function
Error:
On Error GoTo -1
On Error Resume Next
If Requests.RequestCount > 0 Then Requests.RemoveRequest CurrKey
Program.addLog "(2captcha) [" & Replace(Left$(CurrKey, InStr(CurrKey & ",", ",") - 1), "-", vbNullString, , 1) + 1 & "] Error!"
End Function

Private Sub Requests_ResponseFinished(Req As cAsyncRequest)
Dim TmpStr(2) As String, strK As String
TmpStr(2) = Req.Key
If Right$(TmpStr(2), 1) = " " Then
Requests.RemoveRequest TmpStr(2)
On Error GoTo E1
Debug.Print Req.http.ResponseText
If Req.http.ResponseText = "OK_REPORT_RECORDED" Then
Program.addLog "(2captcha) Following captcha text has been reported as incorrect: " & RTrim$(Mid$(TmpStr(2), InStr(TmpStr(2), "-") + 1))
Else
E1: Program.addLog "(2captcha) Reporting of following captcha text as incorrect has been failed: " & RTrim$(Mid$(TmpStr(2), InStr(TmpStr(2), "-") + 1))
End If
Else
Dim bytT(1) As Byte, strErr As String, i As Byte
TmpStr(1) = Req.Tag
On Error GoTo E
If TmpStr(1) = vbNullString Then
RandStr TmpStr(0)
Dim PostData() As Byte
PostData = ""
bytT(0) = InStr(TmpStr(2), ",") + 1
If bytT(0) > 1 Then
For i = 0 To 5
bytT(1) = InStr(bytT(0), TmpStr(2), ",")
If bytT(1) = 0 Then bytT(1) = Len(TmpStr(2)) + 1
If bytT(1) > bytT(0) Then
If i > 0 Then
bytT(0) = Mid$(TmpStr(2), bytT(0), bytT(1) - bytT(0))
If Not frmS.bolNoRem Then
Select Case i
Case 1: frmS.bolCase = bytT(0) = 1
Case 2: frmS.bolPhrase = bytT(0) = 1
Case 3: frmS.bolMath = bytT(0) = 1
Case 4: frmS.bytType = bytT(0)
Case 5: frmS.bytLang = bytT(0)
End Select
End If
If bytT(0) > 0 Then CatBinaryString PostData, "--" & TmpStr(0) & vbCrLf & "Content-Disposition: form-data; name=""" & Prms(i - 1) & """" & vbCrLf & vbCrLf & bytT(0) & vbCrLf
Else
strK = Mid$(TmpStr(2), bytT(0), bytT(1) - bytT(0))
If Not frmS.bolNoRem Then frmS.strKey = strK
End If
End If
bytT(0) = bytT(1) + 1
If bytT(0) >= Len(TmpStr(2)) Then Exit For
Next
End If
Dim CT As String: CT = Req.http.GetResponseHeader("Content-Type")
CatBinaryString PostData, "--" & TmpStr(0) & vbCrLf & "Content-Disposition: form-data; name=""soft_id""" & vbCrLf & vbCrLf & SOFT_KEY & vbCrLf & _
"--" & TmpStr(0) & vbCrLf & "Content-Disposition: form-data; name=""key""" & vbCrLf & vbCrLf & strK & vbCrLf & _
"--" & TmpStr(0) & vbCrLf & "Content-Disposition: form-data; name=""method""" & vbCrLf & vbCrLf & "post" & vbCrLf & _
"--" & TmpStr(0) & vbCrLf & "Content-Disposition: form-data; name=""file""; filename=""captcha." & Mid$(CT, InStr(CT, "/") + 1) & """" & vbCrLf & "Content-Type: " & CT & vbCrLf & vbCrLf
CatBinary PostData, Req.http.ResponseBody
CatBinaryString PostData, vbCrLf & "--" & TmpStr(0) & "--" & vbCrLf
Program.addLog "(2captcha) [" & Mid$(Left$(TmpStr(2), InStr(TmpStr(2) & ",", ",") - 1), 2) + 1 & "] Potential captcha downloaded. Submitting captcha..."
Req.Tag = "+" & bytT(0)
Req.SendRequest "http://" & SERVER & "/in.php", PostData, "Content-Type" & vbLf & "multipart/form-data; boundary=" & TmpStr(0)
Erase PostData
Else
TmpStr(0) = Req.http.ResponseText
bytT(0) = InStr(TmpStr(0), "|")
If bytT(0) = 0 Then
If TmpStr(0) <> "CAPCHA_NOT_READY" Then
Select Case TmpStr(0)
Case "ERROR_CAPTCHA_UNSOLVABLE": strErr = "Captcha could not be solved"
Case "ERROR_KEY_DOES_NOT_EXIST": strErr = "The key doesn't exist"
Case "ERROR_ZERO_BALANCE": strErr = "No money on the account"
Case "ERROR_NO_SLOT_AVAILABLE": strErr = "The current bid is higher than the maximum bid set for the account"
Case "ERROR_ZERO_CAPTCHA_FILESIZE": strErr = "Image size is less than 100 bits"
Case "ERROR_TOO_BIG_CAPTCHA_FILESIZE": strErr = "Image size is more than 100 Kb"
Case "ERROR_IP_NOT_ALLOWED": strErr = "The request was sent from the IP that is not on the list of your IPs"
Case "IP_BANNED": strErr = "The IP address from which is accessed is banned"
End Select
Debug.Print TmpStr(0)
GoTo E
Else
Dim endTime(1) As Double
bytT(0) = InStr(TmpStr(1), "|") + 1
If bytT(0) > 1 Then
bytT(1) = InStr(bytT(0), TmpStr(1), "|")
If bytT(1) = 1 Then bytT(1) = Len(TmpStr(1)) + 1
bytT(0) = Mid$(TmpStr(1), bytT(0), bytT(1) - bytT(0))
If bytT(0) > 0 Then endTime(0) = Timer + bytT(0) Else: endTime(0) = Timer + frmS.bytWait
If bytT(1) < Len(TmpStr(1)) Then endTime(1) = Mid$(TmpStr(1), bytT(1) + 1)
Else: endTime(0) = Timer + frmS.bytWait
End If
End If
ElseIf Left$(TmpStr(1), 1) <> "+" Then
CaptchaText.Add Mid$(TmpStr(0), bytT(0) + 1) & "|" & Left$(TmpStr(1), InStr(TmpStr(1) & "|", "|") - 1), Mid$(TmpStr(2), 2) & "-"
CaptchaText.Remove TmpStr(2)
Exit Sub
Else
bytT(1) = TmpStr(1)
TmpStr(1) = Mid$(TmpStr(0), bytT(0) + 1)
If bytT(1) > 1 And bytT(1) < Len(TmpStr(2)) Then
Dim strT As String, bytT2 As Byte
For i = 0 To 1
bytT2 = InStr(bytT(1), TmpStr(2), ",")
If bytT2 = 0 Then bytT2 = Len(TmpStr(2)) + 1
bytT(1) = Mid$(TmpStr(2), bytT(1), bytT2 - bytT(1))
If bytT(1) > 0 Then endTime(i) = Timer + bytT(1) Else: If i = 0 Then endTime(0) = Timer + frmS.bytWait
If i = 0 Then strT = strT & "|" & bytT(1) Else: If bytT(1) > 0 Then strT = strT & "|" & endTime(1)
If Not frmS.bolNoRem Then If i = 1 Then frmS.bytMax = bytT(1) Else: If bytT(1) > 0 Then frmS.bytWait = bytT(1)
If bytT2 < Len(TmpStr(2)) Then bytT(1) = bytT2 + 1 Else: Exit For
Next
If strT <> "|0" Then Req.Tag = TmpStr(1) & strT Else: Req.Tag = TmpStr(1)
Else
Req.Tag = TmpStr(1)
endTime(0) = Timer + frmS.bytWait
If frmS.bytMax > 0 Then
endTime(1) = Timer + frmS.bytMax
Req.Tag = TmpStr(1) & "|0|" & endTime(1)
End If
End If
End If
Do While Timer < endTime(0)
If endTime(1) > 0 Then If Timer >= endTime(1) Then Exit Do
DoEvents
Loop
If endTime(1) = 0 Or Timer < endTime(1) Then
Program.addLog "(2captcha) [" & Mid$(Left$(TmpStr(2), InStr(TmpStr(2) & ",", ",") - 1), 2) + 1 & "] Checking if captcha is solved..."
strK = Mid$(TmpStr(2), InStr(TmpStr(2), ",") + 1)
Req.SendRequest "http://" & SERVER & "/res.php?key=" & Left$(strK, InStr(strK & ",", ",") - 1) & "&action=get&id=" & Left$(TmpStr(1), InStr(TmpStr(1) & "|", "|") - 1)
Else
strErr = "Response timeout"
GoTo E
End If
End If
Exit Sub
E:
If strErr = vbNullString Then strErr = "!" Else: strErr = ": " & strErr & "."
Program.addLog "(2captcha) [" & Mid$(Left$(TmpStr(2), InStr(TmpStr(2) & ",", ",") - 1), 2) + 1 & "] Error" & strErr
CaptchaText.Remove TmpStr(2)
End If
End Sub

Private Sub Requests_Error(Req As cAsyncRequest, ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
Program.addLog "(2captcha) [" & Mid$(Left$(Req.Key, InStr(Req.Key & ",", ",") - 1), 2) + 1 & "] Error!"
CaptchaText.Remove Req.Key
End Sub

Private Sub Tmr_Timer(ByVal Seconds As Currency)
Do While CaptchaID.Count > 0
CaptchaID.Remove 1
Loop
End Sub
