VERSION 5.00
Begin VB.Form frmSB 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd0 
      Caption         =   "&Y"
      Height          =   255
      Index           =   0
      Left            =   3600
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   120
      Width           =   255
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      ItemData        =   "frmSB.frx":0000
      Left            =   120
      List            =   "frmSB.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Type"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   120
      TabIndex        =   16
      ToolTipText     =   "Input"
      Top             =   2760
      Width           =   3255
   End
   Begin VB.TextBox txtSource 
      Height          =   765
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   15
      ToolTipText     =   "Source"
      Top             =   1920
      Width           =   3255
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "&Test"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3960
      Width           =   1575
   End
   Begin VB.ComboBox cmbCase 
      Height          =   315
      ItemData        =   "frmSB.frx":0035
      Left            =   1440
      List            =   "frmSB.frx":0042
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "Case"
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox cmbEnc 
      Height          =   315
      ItemData        =   "frmSB.frx":005C
      Left            =   2425
      List            =   "frmSB.frx":0069
      Style           =   2  'Dropdown List
      TabIndex        =   5
      ToolTipText     =   "Encryption"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "&Insert"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmdNL 
      Caption         =   "&New line"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   855
   End
   Begin VB.CheckBox chkDigit 
      Caption         =   "&Digits"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   735
   End
   Begin VB.CheckBox chkLetters 
      Caption         =   "L&etters"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1980
      TabIndex        =   12
      Top             =   1200
      Width           =   855
   End
   Begin VB.CheckBox chkSym 
      Caption         =   "&Symbols"
      Enabled         =   0   'False
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   1200
      Width           =   920
   End
   Begin VB.TextBox txtCustom 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1725
      TabIndex        =   7
      Top             =   840
      Width           =   580
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "O&K"
      Height          =   255
      Left            =   1800
      TabIndex        =   19
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox txtOutput 
      Height          =   765
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   17
      ToolTipText     =   "Output"
      Top             =   3120
      Width           =   3255
   End
   Begin VB.TextBox txtLength 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3000
      TabIndex        =   9
      ToolTipText     =   "Leave blank or zero for unlimited."
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtIns 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "&Replace"
      Height          =   255
      Left            =   1800
      TabIndex        =   14
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton cmd0 
      Caption         =   "&R"
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton cmd0 
      Caption         =   "&S"
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmd0 
      Caption         =   "&E"
      Height          =   255
      Index           =   3
      Left            =   3600
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton cmd0 
      Caption         =   "&U"
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton cmd0 
      Caption         =   "&P"
      Height          =   255
      Index           =   5
      Left            =   3600
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton cmd0 
      Caption         =   "&O"
      Height          =   255
      Index           =   8
      Left            =   3600
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox txtRepl 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      ToolTipText     =   "Replacement"
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmd0 
      Caption         =   "&L"
      Height          =   255
      Index           =   9
      Left            =   3600
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label lblL 
      Caption         =   "Lengt&h:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblC 
      Caption         =   "&Custom:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   840
      Width           =   615
   End
End
Attribute VB_Name = "frmSB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strRI As String, bytT As Byte

Private Sub chkLetters_Click()
chkDigit_Click
End Sub

Private Sub chkSym_Click()
chkDigit_Click
End Sub

Private Sub chkDigit_Click()
If cmbType.ListIndex = 2 Then txtIns.Text = "[inp" Else: txtIns.Text = "[rnd"
strRI = vbNullString
If chkDigit.Value = 1 Then
txtIns.Text = txtIns.Text & "D"
strRI = strDigit
End If
If chkLetters.Value = 1 Then
AddLett strRI
If cmbCase.ListIndex = 0 Then txtIns.Text = txtIns.Text & "L"
End If
If chkSym.Value = 1 Then
txtIns.Text = txtIns.Text & "S"
strRI = strRI & strSym
End If
If strRI = vbNullString Then
strRI = strDigit
If cmbType.ListIndex = 3 Then AddLett strRI
strRI = strRI & strSym
AddLett strRI
End If
Dim strT As String
If txtCustom.Text <> vbNullString Then
Dim i As Byte
For i = 1 To Len(txtCustom.Text)
If InStr(strT & strRI, Mid$(txtCustom.Text, i, 1)) = 0 Then strT = strT & Mid$(txtCustom.Text, i, 1)
Next
If strT <> vbNullString Then
txtIns.Text = txtIns.Text & "`" & Replace(strT, "`", "``") & "`"
strRI = strRI & txtCustom.Text
End If
End If
strT = vbNullString
On Error Resume Next
If cmbType.ListIndex = 3 Then
strT = CInt(txtLength.Text)
Else
strT = txtLength.Text
If InStr(strT, "-") > 0 Then If CInt(Split(strT, "-")(0)) = 0 Then strT = Split(strT, "-")(1) Else: If CInt(Split(strT, "-")(0)) > CInt(Split(strT, "-")(1)) And CInt(Split(strT, "-")(1)) > 0 Then strT = vbNullString
End If
If strT <> vbNullString Then If InStr(strT, "-") = 0 Then If strT = 0 Then strT = vbNullString
txtIns.Text = txtIns.Text & strT & "]"
End Sub

Private Function AddLett(strRI As String)
If cmbType.ListIndex = 3 Then
Select Case cmbCase.ListIndex
Case 0: strRI = strRI & strLett
Case 1
strRI = strRI & strULett
txtIns.Text = txtIns.Text & "U"
Case 2
strRI = strRI & strLett & strULett
txtIns.Text = txtIns.Text & "M"
End Select
Else: strRI = strRI & strLett & strULett
End If
End Function

Private Sub cmbCase_Click()
If cmbType.ListIndex < 3 Then Exit Sub
chkDigit_Click
If cmbType.ListIndex <> 3 Then Exit Sub
If cmbCase.ListIndex > 0 Then
chkLetters.Value = 1
chkLetters.Enabled = False
Else
chkLetters.Enabled = True
chkLetters.Value = 0
End If
End Sub

Private Sub cmd0_Click(Index As Integer)
Select Case Index
Case 0: cmbType.SetFocus
Case 1: If txtIns.Enabled Then txtIns.SetFocus
Case 2: cmbCase.SetFocus
Case 3: cmbEnc.SetFocus
Case 4: txtSource.SetFocus
Case 5: txtInput.SetFocus
Case 6: txtOutput.SetFocus
Case 7: If cmbType.ListIndex = 1 Then txtRepl.SetFocus
End Select
End Sub

Private Sub cmdInsert_Click()
If txtIns.Text = vbNullString Then Exit Sub
Dim strT As String, bolS(1) As Boolean
PrepR strT
If cmbType.ListIndex <> 1 Then
If txtInput.SelStart > 0 Then
If txtInput.SelText <> vbNullString Then
If txtInput.SelStart < Len(txtInput.Text) - 1 Then
DetR bolS(0), bolS(1)
If bolS(0) Or bolS(1) Then
If bolS(1) Then If Mid$(txtInput.Text, txtInput.SelStart + 1, 1) = "'" Then strT = "'+" & strT
If bolS(0) Then If Mid$(txtInput.Text, txtInput.SelStart + txtInput.SelLength + 1, 1) = "'" Then strT = strT & "+'"
End If
Else: strT = "+" & strT
End If
ElseIf txtInput.SelStart = Len(txtInput.Text) Then If Right$(txtInput.Text, 1) <> "+" Then strT = "+" & strT
End If
Else: strT = strT & "+"
End If
Else: AddR strT
End If
BackTo strT
cmbType.SetFocus
End Sub

Private Sub cmdOK_Click()
Do While Left$(txtInput.Text, 1) = "+"
txtInput.Text = Mid$(txtInput.Text, 2)
Loop
Do While Right$(txtInput.Text, 1) = "+"
txtInput.Text = Left$(txtInput.Text, Len(txtInput.Text) - 1)
Loop
frmMain.txtExp(Me.Tag).Text = txtInput.Text
Unload Me
End Sub

Sub BackTo(strT As String, Optional bolT As Boolean)
Dim intS As Integer
intS = txtInput.SelStart
txtInput.SelText = strT
If Not bolT Then txtInput.SelStart = intS + Len(strT) Else: txtInput.SelStart = intS + InStrRev(strT, ",")
If cmbType.ListIndex < 2 Then txtIns.Text = vbNullString
If cmbType.ListIndex <> 3 Then
cmbCase.ListIndex = 0
If cmbType.ListIndex = 1 Then txtRepl.Text = vbNullString
End If
cmbEnc.ListIndex = 0
End Sub

Private Sub cmdReplace_Click()
If txtIns.Text = vbNullString Then Exit Sub
Dim strT As String
PrepR strT
AddR strT, True
BackTo strT, True
cmd0_Click 1
End Sub

Private Function PrepR(strT As String)
If strT = vbNullString Then
strT = "'" & Rpl(txtIns.Text) & "'"
If cmbType.ListIndex <> 1 Then
If cmbType.ListIndex <> 3 Then
If cmbType.ListIndex = 0 Then
strT = Replace(strT, "[inp", "['+'inp")
strT = Replace(strT, "[rnd", "['+'rnd")
End If
End If
Else
If txtRepl.Text <> vbNullString Then strT = strT & ",'" & Rpl(txtRepl.Text) & "'"
Exit Function
End If
End If
If cmbType.ListIndex <> 3 And cmbCase.ListIndex > 0 Then If cmbCase.ListIndex = 1 Then strT = "l(" & strT & ")" Else: strT = "u(" & strT & ")"
If cmbEnc.ListIndex = 1 Then strT = "b64(" & strT & ")" Else: If cmbEnc.ListIndex = 2 Then strT = "md5(" & strT & ")"
End Function

Private Function AddR(strT As String, Optional bolRpl As Boolean)
On Error Resume Next
Dim strS(1) As String, bolS(1) As Boolean
If Not bolRpl Then
strS(0) = "rg"
Dim strT1 As String
If IsNumeric(Replace(txtCustom.Text, ",", vbNullString)) Then
strT1 = "," & CByte(Split(txtCustom.Text & ",", ",")(1))
If strT1 = ",0" Then strT1 = vbNullString
strT1 = "," & CByte(Split(txtCustom.Text & ",", ",")(0)) & strT1
End If
strS(1) = strT1 & ")"
Else
strS(0) = "rpl"
strS(1) = ",)"
End If
If txtInput.SelText <> vbNullString Then
strT = strS(0) & "(" & DetR(bolS(0), bolS(1)) & "," & strT & strS(1)
PrepR strT
If bolS(1) Then If Mid$(txtInput.Text, txtInput.SelStart + 1, 1) = "'" Then strT = "'+" & strT
If bolS(0) Then If Mid$(txtInput.Text, txtInput.SelStart + txtInput.SelLength + 1, 1) = "'" Then strT = strT & "+'"
Else
strT = strS(0) & "('[src]'," & strT & strS(1)
PrepR strT
If txtInput.SelStart = 0 Then strT = strT & "+" Else: If txtInput.SelStart = Len(txtInput.Text) Then If Right$(txtInput.Text, 1) <> "+" Then strT = "+" & strT
End If
End Function

Private Function DetR(bolS0 As Boolean, bolS1 As Boolean) As String
DetR = txtInput.SelText
If UBound(Split(DetR, "'")) > 0 Then
If Right$(DetR, 1) <> "'" Or Mid$(DetR, Len(DetR) - 1, 1) = "'" Then
DetR = Rpl(DetR) & "'"
bolS0 = True
End If
If Left$(DetR, 1) <> "'" Or Mid$(DetR, Len(Split("(" & DetR, "(")(1)) + 1, 1) = "'" Then
If bolS0 Then DetR = "'" & DetR Else: DetR = "'" & Rpl(DetR)
bolS1 = True
End If
Else
DetR = "'" & Rpl(DetR) & "'"
bolS0 = True
bolS1 = True
End If
End Function

Private Sub txtIns_GotFocus()
If txtInput.SelText = vbNullString Then Exit Sub
Dim strT As String
strT = txtInput.SelText
If Left$(strT, 1) = "'" Then strT = Mid$(strT, 2)
If Right$(strT, 1) = "'" Then strT = Left$(strT, Len(strT) - 1)
txtIns.Text = strT
txtIns.SelLength = Len(txtIns.Text)
End Sub

Private Sub cmbType_Click()
If bytT = cmbType.ListIndex Then Exit Sub
If cmbType.ListIndex = 0 Then cmdNL.Enabled = True Else: cmdNL.Enabled = False
If bytT < 2 Then
If txtIns.Text <> vbNullString Then
If MsgBox("Sure?", vbQuestion + vbYesNo) = vbNo Then
cmbType.ListIndex = bytT
txtIns.SetFocus
Exit Sub
End If
End If
End If
If cmbType.ListIndex <> 3 And bytT = 3 Then
cmbCase.RemoveItem 2
cmbCase.AddItem "Normal", 0
cmbCase.ListIndex = 0
End If
If cmbType.ListIndex > 0 Then
lblC.Enabled = True
txtCustom.Enabled = True
Else
lblC.Enabled = False
txtCustom.Enabled = False
End If
If cmbType.ListIndex = 1 Then
txtRepl.Enabled = True
txtIns.Width = 1575
ElseIf bytT = 1 Then
txtIns.Width = 3255
txtRepl.Enabled = False
End If
If cmbType.ListIndex < 2 Then
txtIns.Text = vbNullString
txtIns.Enabled = True
chkDigit.Enabled = False
chkLetters.Enabled = False
chkSym.Enabled = False
lblL.Enabled = False
txtLength.Enabled = False
Else
txtIns.Enabled = False
If cmbType.ListIndex = 3 Then
txtIns.Text = "[rnd]"
cmbCase.RemoveItem 0
cmbCase.AddItem "Mixed"
cmbCase.ListIndex = 0
Else: txtIns.Text = "[inp]"
End If
chkDigit.Enabled = True
chkLetters.Enabled = True
chkSym.Enabled = True
lblL.Enabled = True
txtLength.Enabled = True
chkDigit_Click
End If
bytT = cmbType.ListIndex
End Sub

Private Sub cmdNL_Click()
txtIns.Text = txtIns.Text & "[nl]"
txtIns.SetFocus
txtIns.SelStart = Len(txtIns.Text)
End Sub

Private Sub Form_Load()
If frmMain.chkOnTop.Checked Then SetTopMostWindow Me.hWnd, True
cmbType.ListIndex = 0
cmbCase.ListIndex = 0
cmbEnc.ListIndex = 0
'del
'Dim obj As IPluginInterface
'Set plug = CreateObjectEx2("C:\Documents and Settings\MikiSoft\My Documents\Dropbox\UniBot\Plugins\SimpleCaptcha.dll", "C:\Documents and Settings\MikiSoft\My Documents\Dropbox\UniBot\Plugins\SimpleCaptcha.dll", "SimpleCaptcha")
'plug.Startup Me
'Plugins.add plug
'Set plug = Nothing
'txtSource.Text = "<center><h1>Bot Protection</h1></br>" & vbCrLf & "<img src=""http://smashbtc.com/members/avatar/default_11.jpg"" style=""width:100px;height:100px;""/></br>" & vbCrLf & "Click on same picture :</br>" & vbCrLf & "<a href=""index.php?view=account&ac=btc-collect&action=solve&cid=17&sid=10TVM0eU9USXhOelF5Tn&sid2=10TVM&siduid=10&""><img src=""http://smashbtc.com/members/avatar/default_14.jpg"" style=""image-orientation: 180deg;" & _
" width:100px;height:100px;""/></a>" & vbCrLf & "<a href=""index.php?view=account&ac=btc-collect&action=solve&cid=9&sid=10TVM0eU9USXhOelF5Tn&sid2=10TVM&siduid=10&""><img src=""http://smashbtc.com/members/avatar/default_4.jpg"" style=""image-orientation: 180deg; width:100px;height:100px;""/></a>" & vbCrLf & "" & vbCrLf & "<a href=""index.php?view=account&ac=btc-collect&action=solve&cid=11&sid=10TVM0eU9USXhOelF5Tn&sid2=10TVM" & _
"&siduid=10&""><img src=""http://smashbtc.com/members/avatar/default_11.jpg"" style=""image-orientation: 180deg; width:100px;height:100px;""/></a>" & vbCrLf & "" & vbCrLf & "<a href=""index.php?view=account&ac=btc-collect&action=solve&cid=17&sid=10TVM0eU9USXhOelF5Tn&sid2=10TVM&siduid=10&""><img src=""http://smashbtc.com/members/avatar/default_16.jpg"" style=""image-orientation: 180deg; width:100px;height:100px;""/></a>"
'txtInput.Text = "'http://smashbtc.com/'+rg('[src]','href=""(.*?)""><img src=""'+rg('[src]','Bot Protection<\/h1><\/br>\r\n<img src=""(.*?)""','$1')+'""','$1')"
'cmdTest_Click
'del
End Sub

'del
'Private Sub Form_Unload(Cancel As Integer)
'On Error Resume Next
'Plugins.Remove 1
'UnloadLibrary "C:\Documents and Settings\MikiSoft\My Documents\Dropbox\UniBot\Plugins\SimpleCaptcha.dll"
'End Sub
'del

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me: Exit Sub
If ActiveControl Is Nothing Then Exit Sub
Const ASC_CTRL_A As Integer = 1

    ' See if this is Ctrl-A.
    If KeyAscii = ASC_CTRL_A Then
    KeyAscii = 0
        ' The user is pressing Ctrl-A. See if the
        ' active control is a TextBox.
        If TypeOf ActiveControl Is TextBox Then
            ' Select the text in this control.
            SelectT ActiveControl, 0, Len(ActiveControl.Text)
        End If
    End If
End Sub

Private Sub cmdTest_Click()
Dim strOut As String, intCount As Integer, strT As String
strT = ReplaceString(txtInput.Text)
strOut = Replace(ProcessString(strT, txtSource.Text, , , intCount), vbNewLine, "[nl]")
If intCount > 1 Then
Dim i As Integer, intStart As Integer
For i = 1 To intCount - 1
intStart = i
strOut = strOut & vbNewLine & Replace(ProcessString(strT, txtSource.Text, , intStart), vbNewLine, "[nl]")
Next
End If
txtOutput.Text = strOut
'txtOutput.SetFocus 'enable
End Sub

Private Sub txtCustom_Change()
If cmbType.ListIndex = 3 Then
If UBound(Split(txtCustom.Text & "-", "-")) = 2 Then
If IsNumeric(Replace(txtCustom.Text, "-", vbNullString)) And Not InStr(txtCustom.Text, ",") > 0 And Not InStr(txtCustom.Text, "d") > 0 Then
Dim intNum(1) As Integer
On Error GoTo E
intNum(0) = CInt("0" & Split(txtCustom.Text, "-")(0))
intNum(1) = CInt("0" & Split(txtCustom.Text, "-")(1))
If intNum(0) < intNum(1) Then
If cmbType.ListIndex = 2 Then txtIns.Text = "[inp" & intNum(0) & "-" & intNum(1) & "]" Else: txtIns.Text = "[rnd" & intNum(0) & "-" & intNum(1) & "]"
Exit Sub
End If
End If
End If
ElseIf cmbType.ListIndex = 1 Then Exit Sub
End If
E: chkDigit_Click
End Sub

Private Sub txtLength_Change()
If cmbType.ListIndex = 3 Then CheckText txtLength Else: If InStr("0123456789-", Right$(txtLength.Text, 1)) = 0 Then Exit Sub
chkDigit_Click
End Sub

Private Function Rpl(strIns As String) As String
Rpl = Replace(Replace(strIns, "'", "''"), vbNewLine, "[nl]")
End Function
