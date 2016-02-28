VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SimpleCaptcha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IPluginInterface
Dim Program As Form, bytI As Byte

Private Sub Class_Terminate()
If Not frmSettings.bolNoCurr Or Program Is Nothing Then Exit Sub 'Exit immediately if there is option enabled not to save settings or if there was some problem while loading plugin
Dim strT As String: IPluginInterface_BuildSettings strT
On Error Resume Next
If strT <> vbNullString Then
Open App.Path & "\" & App.EXEName & ".ini" For Output Access Write As #1 'Write to INI settings file (MUST be file with that extension!)
Print #1, strT;
Close #1
If bolHid Then SetAttr App.Path & "\" & App.EXEName & ".ini", vbHidden
Else: Kill App.Path & "\" & App.EXEName & ".ini"
End If
End Sub

Private Sub IPluginInterface_BuildSettings(strS As String)
If Not frmSettings.bolNoLett And Not frmSettings.bolNoDigits And Not frmSettings.bolNoSym And Not frmSettings.strCustom <> vbNullString And frmSettings.bytCase = 0 And frmSettings.bytMin = 3 And frmSettings.bytMax = 0 And frmSettings.bytTimeout = 10 And (Not frmSettings.bolOCR Or frmSettings.bolOCR = vbUseDefault) And Not frmSettings.bolNoRem Then Exit Sub 'Check if there are changes in settings
If frmSettings.bytMin <> 3 Then strS = "min=" & frmSettings.bytMin & vbCrLf
If frmSettings.bytMax > 0 Then strS = strS & "max=" & frmSettings.bytMax & vbCrLf
If frmSettings.bytTimeout <> 10 Then strS = strS & "timeout=" & frmSettings.bytTimeout & vbCrLf
If frmSettings.bolNoLett Then strS = strS & "noletters=1" & vbCrLf
If frmSettings.bolNoDigits Then strS = strS & "nodigits=1" & vbCrLf
If frmSettings.bolNoSym Then strS = strS & "nosymbols=1" & vbCrLf
If frmSettings.strCustom <> vbNullString Then strS = strS & "custom=" & frmSettings.strCustom & vbCrLf
If Not frmSettings.bolNoLett Then If frmSettings.bytCase > 0 Then strS = strS & "case=" & frmSettings.bytCase & vbCrLf
If frmSettings.bolOCR Then strS = strS & "ocr=1" & vbCrLf
If frmSettings.bolNoRem Then strS = strS & "noremember=1" & vbCrLf
strS = Left$(strS, Len(strS) - 2)
End Sub

Private Function IPluginInterface_Startup(Parent As Object) As String
Set Program = Parent
frmSettings.bytMin = 3
frmSettings.bytTimeout = 10
On Error Resume Next
Dim objWSH As Object
Set objWSH = CreateObject("WScript.Shell")
If InStr(1, objWSH.Environment("USER")("PATH"), "tesseract", vbTextCompare) = 0 And InStr(1, objWSH.Environment("SYSTEM")("PATH"), "tesseract", vbTextCompare) = 0 Then frmSettings.bolOCR = vbUseDefault 'Check if there's no Tesseract OCR present on system
Set objWSH = Nothing
On Error GoTo E
If Dir$(App.Path & "\" & App.EXEName & ".ini", vbHidden) <> vbNullString Then 'Check if settings are present (file name MUST BE the same like DLL and must also have check if it's HIDDEN!)
bolHid = GetAttr(App.Path & "\" & App.EXEName & ".ini") = vbHidden 'Add flag if file is hidden
Dim strL(1) As String
Open App.Path & "\" & App.EXEName & ".ini" For Input Access Read As #1 'Load settings file
While Not EOF(1)
Line Input #1, strL(0)
strL(0) = Trim$(Replace(strL(0), vbCr, vbNullString))
If strL(0) <> vbNullString Then
If InStr(";#[", Left$(strL(0), 1)) = 0 Then
strL(1) = Left$(strL(0), InStr(strL(0), "=") - 1)
strL(0) = Mid$(Left$(strL(0), Len(strL(0))), Len(strL(1)) + 2)
If IsNumeric(strL(0)) Then
Select Case strL(1)
Case "min": If strL(0) < 255 Then frmSettings.bytMin = strL(0) Else: frmSettings.bytMin = 255
Case "max": If strL(0) < 255 Then frmSettings.bytMax = strL(0) Else: frmSettings.bytMax = 255
Case "timeout": If strL(0) < 255 Then frmSettings.bytTimeout = strL(0) Else: frmSettings.bytTimeout = 255
Case "ocr": If frmSettings.bolOCR = vbFalse Then frmSettings.bolOCR = strL(0) > 0
Case "noletters": If frmSettings.strCustom <> vbNullString Or (frmSettings.bolNoDigits Xor frmSettings.bolNoSym) Or Not frmSettings.bolNoDigits And Not frmSettings.bolNoSym Then frmSettings.bolNoLett = strL(0) > 0
Case "nodigits": If frmSettings.strCustom <> vbNullString Or (frmSettings.bolNoLett Xor frmSettings.bolNoSym) Or Not frmSettings.bolNoLett And Not frmSettings.bolNoSym Then frmSettings.bolNoDigits = strL(0) > 0
Case "nosymbols": If frmSettings.strCustom <> vbNullString Or (frmSettings.bolNoDigits Xor frmSettings.bolNoLett) Or Not frmSettings.bolNoDigits And Not frmSettings.bolNoLett Then frmSettings.bolNoSym = strL(0) > 0
Case "custom": frmSettings.strCustom = strL(0)
Case "case": If strL(0) = "1" Then frmSettings.bytCase = 1 Else: If strL(0) = "2" Then frmSettings.bytCase = 2
Case "noremember": frmSettings.bolNoRem = strL(0) > 0
End Select
End If
End If
End If
Wend
Close #1
If frmSettings.bytMax > 0 And frmSettings.bytMax < frmSettings.bytMin Then frmSettings.bytMax = frmSettings.bytMin
If frmSettings.strCustom <> vbNullString Then frmSettings.strCustom = ProcStr(frmSettings.strCustom, frmSettings.bolNoLett, frmSettings.bolNoDigits, frmSettings.bolNoSym, frmSettings.bytCase)
If frmSettings.strCustom = vbNullString And frmSettings.bolNoLett And frmSettings.bolNoDigits And frmSettings.bolNoSym Then
frmSettings.bolNoLett = False
frmSettings.bolNoDigits = False
frmSettings.bolNoSym = False
End If
frmSettings.bolNoCurr = True
End If
E:
IPluginInterface_Startup = "captcha" 'Command(s) for calling (separated by commas if they're more than one)
End Function

Private Function IPluginInterface_Info(Inf() As String) As Object
Inf(0) = App.Major & "." & App.Minor 'Version
Inf(1) = "MikiSoft" 'Author
Inf(2) = "Parses strings as URLs to captcha image and displays it to user for solving." & vbNewLine & _
"Example: captcha('http://site.com/captchaimage','useragent','cookiedata','referer',10,3,0,0,'',1,1,1,1)." & vbNewLine & _
"Parameters after referer: Response timeout, Min. & max. length, Case (0 - both, 1 - upper, 2 - lower), Custom characters; Letters, Digits, Symbols, Tesseract OCR (0-1)." & vbNewLine & _
"All parameters after URL are optional." 'Description
Set IPluginInterface_Info = frmSettings 'Return settings form to the program(!)
End Function

Private Function IPluginInterface_Execute(Params() As String) As String 'Params(0) is command which is called (in this case it will be "captcha" without quotes)
Dim Req As WinHttp.WinHttpRequest, strI As String, i As Byte
bytI = bytI + 1
strI = "-" & bytI
'HTTPS settings
Set Req = New WinHttp.WinHttpRequest
Req.Option(4) = 13056
Req.Option(12) = True
On Error GoTo Error 'It's VERY IMPORTANT TO have error handlers in function, in order to PREVENT FATAL ERRORS in program if something goes wrong
Req.Open "GET", Params(1) 'Open GET request with captcha URL (index 0 is command name)
'Set parameters
Dim bytT(3) As Byte, bolT(4) As Boolean, strT As String
strT = frmSettings.strCustom
bytT(0) = frmSettings.bytTimeout
bytT(1) = frmSettings.bytMin
bytT(2) = frmSettings.bytMax
bytT(3) = frmSettings.bytCase
bolT(0) = frmSettings.bolNoLett
bolT(1) = frmSettings.bolNoDigits
bolT(2) = frmSettings.bolNoSym
If frmSettings.bolOCR <> vbUseDefault Then bolT(3) = frmSettings.bolOCR
If UBound(Params) >= 2 Then
If Not IsNumeric(Params(2)) Then
If Params(2) <> vbNullString Then Req.SetRequestHeader "User-Agent", Params(2)
If UBound(Params) >= 3 Then
If Not IsNumeric(Params(3)) Then
If Params(3) <> vbNullString Then Req.SetRequestHeader "Cookie", Params(3)
If UBound(Params) >= 4 Then
If Not IsNumeric(Params(4)) Then
If Params(4) <> vbNullString Then Req.SetRequestHeader "Referer", Params(4)
If UBound(Params) >= 5 Then i = 5
Else: i = 4
End If
End If
Else: i = 3
End If
End If
Else: i = 2
End If
If i > 0 Then
Dim a As Byte
For i = i To i + 2
If UBound(Params) < i Then GoTo C
If IsNumeric(Params(i)) Then
If Params(i) < 0 Then Params(i) = Params(i) * (-1)
If Params(i) < 255 Then bytT(a) = Params(i) Else: bytT(a) = 255
If Not frmSettings.bolNoRem Then
Select Case a
Case 0: frmSettings.bytTimeout = bytT(0)
Case 1: frmSettings.bytMin = bytT(1)
Case 2: frmSettings.bytMax = bytT(2)
End Select
End If
End If
a = a + 1
Next
If bytT(2) > 0 And bytT(2) < bytT(1) Then bytT(2) = bytT(1)
If UBound(Params) >= i Then
If Params(i) = "0" Or Params(i) = "1" Or Params(i) = "2" Then
bytT(3) = Params(i)
If Not frmSettings.bolNoRem Then frmSettings.bytCase = bytT(3)
bytT(3) = bytT(3) + 1
End If
i = i + 1
If UBound(Params) >= i Then
If Not IsNumeric(Params(i)) Then
strT = Params(i)
If Not frmSettings.bolNoRem Then frmSettings.strCustom = strT
i = i + 1
If UBound(Params) < i Then GoTo C
End If
a = 0
For i = i To i + 2
If UBound(Params) < i Then GoTo C
If Params(i) = "0" Or Params(i) = "1" Then
bolT(a) = CBool(Params(i))
If Not frmSettings.bolNoRem Then
Select Case a
Case 0: frmSettings.bolNoLett = bolT(0)
Case 1: frmSettings.bolNoDigits = bolT(1)
Case 2: frmSettings.bolNoSym = bolT(2)
End Select
End If
End If
a = a + 1
Next
If strT <> vbNullString Then strT = ProcStr(strT, bolT(0), bolT(1), bolT(2), bytT(3))
If strT = vbNullString And bolT(0) And bolT(1) And bolT(2) Then
bolT(0) = False
bolT(1) = False
bolT(2) = False
End If
If frmSettings.bolOCR <> vbUseDefault And UBound(Params) >= i Then
If Params(i) = "0" Or Params(i) = "1" Then
bolT(3) = CBool(Params(i))
If Not frmSettings.bolNoRem Then frmSettings.bolOCR = bolT(3)
End If
End If
End If
End If
End If
End If
C:
Req.Send 'Send request
'Save image as temporary file
Open "tmp" & App.ThreadID & bytI & ".img" For Output Access Write As #1
Print #1, FromCPString(Req.ResponseBody) 'Convert to appropriate format and put received data into file
Close #1
If bolHid Then SetAttr "tmp" & App.ThreadID & bytI & ".img", vbHidden
bolT(4) = Req.GetResponseHeader("Content-Type") = "image/gif" 'Image is GIF or not?
Set Req = Nothing
Program.addLog "(Simple Captcha) [" & bytI & "] Potential captcha downloaded." 'Add to log what's happening
Dim Status As String
Do While DoEvents = 1 'Wait while current captcha form is still displayed
Loop
frmCaptcha.strR = PrepStr(bolT(0), bolT(1), bolT(2))
frmCaptcha.strR = frmCaptcha.strR & strT
frmCaptcha.T = bytT(0)
frmCaptcha.bytM = bytT(1)
frmCaptcha.bytM1 = bytT(2)
frmCaptcha.bytC = bytT(3)
frmCaptcha.Tess = bolT(3)
frmCaptcha.IsGIF = bolT(4)
frmCaptcha.strI = strI
If Program.lblStatus.Caption <> "Waiting for captcha input." Then
Status = Program.lblStatus.Caption 'Memorize current program status
Program.lblStatus.Caption = "Waiting for captcha input." 'Change status in program
End If
frmCaptcha.Show vbModal, Program 'Display captcha solving window and wait for it to close
Dim CT As String: CT = CaptchaText.Item(strI)
If CT = vbNullChar Then 'Check if image was loaded correctly
Error:
Set Req = Nothing
Program.addLog "(Simple Captcha) [" & bytI & "] Error!"
Else: IPluginInterface_Execute = CT 'Put captcha text into returning string
End If
If Status <> vbNullString Then Program.lblStatus.Caption = Status 'Revert back status
On Error GoTo -1
On Error Resume Next
CaptchaText.Remove strI
bytI = bytI - 1
End Function
