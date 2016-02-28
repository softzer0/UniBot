VERSION 5.00
Begin VB.Form frmInput 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkOnly 
      Caption         =   "&Use only on this thread"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "O&K"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtInput 
      Height          =   1335
      Index           =   0
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblName 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1

Dim objI As TextBox, s() As String
Public strInf As String

Private Sub cmdLoad_Click()
Dim strFile As String
strFile = CommDlg(, "Select file to load", "Text file (*.txt)|*.txt|Any file|*.*")
If strFile = vbNullString Then Exit Sub
Dim strT As String
If Trim$(Replace(Replace(objI.Text, vbCr, vbNullString), vbLf, vbNullString)) <> vbNullString Then
If InStr(objI.Text, vbLf) > 0 Then
If Split(objI.Text, vbLf)(UBound(Split(objI.Text, vbLf))) <> vbNullString Then GoTo C
Else
C: strT = vbNewLine
End If
Else: objI.Text = vbNullString
End If
LoadBig objI, strFile, , objI.Text & strT
End Sub

Private Sub cmdOK_Click()
objI.Text = Replace(objI.Text, vbCr, vbNullString)
If Left$(s(0), 1) = "3" Then
frmMain.lblStatus.Caption = "Removing blank lines..."
Screen.MousePointer = 11
frmMain.lblStatus.Refresh
Do While InStr(objI.Text, vbLf & vbLf) > 0
objI.Text = Replace(objI.Text, vbLf & vbLf, vbLf)
Loop
Do While Left$(objI.Text, 1) = vbLf
objI.Text = Mid$(objI.Text, 2)
Loop
Do While Right$(objI.Text, 1) = vbLf
objI.Text = Left$(objI.Text, Len(objI.Text) - 1)
Loop
Screen.MousePointer = 0
End If
If s(2) <> vbNullString Then
frmMain.lblStatus.Caption = "Populating string with input data..."
Screen.MousePointer = 11
frmMain.lblStatus.Refresh
Dim s1() As String
s1() = Split(objI.Text, vbLf)
Dim i As Integer, a As Long, intM(1) As Integer
If UBound(s) = 3 Then
If InStr(s(3), "-") > 0 Then
intM(0) = Val(Split(s(3), "-")(0))
intM(1) = Val(Split(s(3), "-")(1))
Else: intM(1) = s(3)
End If
End If
strInf = vbNullString
For i = 0 To UBound(s1())
For a = 1 To Len(s1(i))
If InStr(s(2), Mid$(s1(i), a, 1)) = 0 Then GoTo N
Next
If intM(0) > 0 Then If Len(s1(i)) < intM(0) Then GoTo N
If intM(1) > 0 Then If Len(s1(i)) > intM(1) Then s1(i) = Left$(s1(i), intM(1))
strInf = strInf & s1(i) & vbLf
N:
Next
Screen.MousePointer = 0
If strInf = vbNullString Then
frmMain.lblStatus.Caption = "..."
objI.Text = vbNullString
If Left$(s(0), 1) <> "3" Then If MsgBox("Input is empty, continue?", vbExclamation + vbYesNo) = vbNo Then Exit Sub
Else
If frmMain.bolDebug Then If txtInput(0).Visible Then frmMain.addLog Replace(Replace(Left$(Me.Caption, InStr(Me.Caption & "; num", "; num") - 1), "Thr", "{T"), "ind", "I") & "} " & UBound(Split(Replace(strInf, vbCr, vbNullString), vbLf)) & " item(s) populated into " & Mid$(s(0), 2) & "." Else: frmMain.addLog Replace(Replace(Left$(Me.Caption, InStr(Me.Caption & "; num", "; num") - 1), "Thr", "{T"), "ind", "I") & "} " & Mid$(s(0), 2) & " [inp]: " & Left$(strInf, Len(strInf) - 1)
strInf = chkOnly.Value & Left$(strInf, Len(strInf) - 1)
End If
Else
If frmMain.bolDebug Then If txtInput(0).Visible Then frmMain.addLog Replace(Replace(Left$(Me.Caption, InStr(Me.Caption & "; num", "; num") - 1), "Thr", "{T"), "ind", "I") & "} " & UBound(Split(Replace(objI.Text, vbCr, vbNullString), vbLf)) + 1 & " item(s) populated into " & Mid$(s(0), 2) & "." Else: frmMain.addLog Replace(Replace(Left$(Me.Caption, InStr(Me.Caption & "; num", "; num") - 1), "Thr", "{T"), "ind", "I") & "} " & Mid$(s(0), 2) & " [inp]: " & objI.Text
strInf = chkOnly.Value & objI.Text
End If
Unload Me
End Sub

Private Sub Form_Load()
If frmMain.chkOnTop.Checked Then SetTopMostWindow Me.hWnd, True
If Left$(strInf, 1) = " " Then
strInf = Mid$(strInf, 2)
chkOnly.Value = 1
chkOnly.Enabled = False
End If
If Left$(strInf, 1) = "0" Or Left$(strInf, 1) = "2" Then
cmdLoad.Enabled = False
txtInput(0).Visible = False
Set objI = txtInput(1)
objI.Visible = True
chkOnly.Top = 720
cmdOK.Top = 1080
cmdOK.Left = 2640
cmdOK.Default = True
Me.Height = 1845
Me.Width = 4065
lblName.Width = 3735
Else: Set objI = txtInput(0)
End If
s() = Split(strInf, vbLf)
strInf = vbNullString
lblName.Caption = s(0)
lblName.Caption = Mid$(lblName.Caption, 2)
If s(2) <> vbNullString Then lblName.ToolTipText = "Allowed chars: " & s(2)
If UBound(s()) = 3 Then
lblName.Caption = lblName.Caption & " (" & s(3) & ")"
If InStr(s(3), "-") > 0 Then cmdOK.Enabled = False
If objI Is txtInput(1) Then
Dim strT As String
If InStr(s(3), "-") > 0 Then strT = Val(Split(s(3), "-")(1)) Else: strT = s(3)
If strT > 0 Then objI.MaxLength = strT
End If
End If
Me.Caption = "Thr: " & Split(s(1), ",")(0) & "; ind: " & Split(s(1), ",")(1)
If Split(s(1), ",")(2) > 0 Then Me.Caption = Me.Caption & "; num: " & Split(s(1), ",")(2)
'txtInput(0).Text = "http://mikisoft.me/programs/unibot/" & vbCrLf & "http://mikisoft.me/programs/uniclicker/" '"http://exc.10khits.com/surf?id=130884&token=0cc44d2df64f0eb464b9355c53d602e1" & vbCrLf & "http://exc.10khits.com/surf?id=130882&token=f8ccfc7abd7a763518702101cf88c9ac" & vbCrLf & "http://exc.10khits.com/surf?id=129032&token=12ed1880591a5bac389c010bfb892738" & vbCrLf & "http://exc.10khits.com/surf?id=130885&token=b82c66da4de717c859cf75bdff182257" 'del
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then strInf = vbNullString
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set objI = Nothing
End Sub

Private Sub txtInput_Change(Index As Integer)
If UBound(s()) < 3 Then Exit Sub
If InStr(s(3), "-") = 0 Then Exit Sub
If Index = 0 Then
If objI.SelLength = 0 Then
Dim lLen As Long: lLen = SendMessage(objI.hWnd, EM_LINELENGTH, EM_LINEINDEX, 0&) - objI.SelLength
If lLen < CInt(Split(s(3), "-")(0)) Then cmdOK.Enabled = False Else: cmdOK.Enabled = True
End If
Else: If Len(objI.Text) < CInt(Split(s(3), "-")(0)) Then cmdOK.Enabled = False Else: cmdOK.Enabled = True
End If
End Sub

Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 22 Or KeyAscii = 3 Or KeyAscii = 19 Or KeyAscii = 24 Then Exit Sub
If Index = 0 Then
If KeyAscii = 13 Or UBound(s()) = 3 Then
If UBound(s()) = 3 And objI.SelLength = 0 Then
Dim lLen As Long: lLen = SendMessage(objI.hWnd, EM_LINELENGTH, EM_LINEINDEX, 0&)
If KeyAscii <> 13 Then
Dim strT As String
If InStr(s(3), "-") > 0 Then strT = Val(Split(s(3), "-")(1)) Else: strT = s(3)
If strT > 0 Then If lLen >= strT Then GoTo E
Else: If InStr(s(3), "-") > 0 Then If lLen < CInt(Split(s(3), "-")(0)) Then GoTo E
End If
End If
If KeyAscii = 13 Then Exit Sub
End If
End If
If s(2) = vbNullString Then Exit Sub
If InStr(s(2), Chr$(KeyAscii)) > 0 Then Exit Sub
E: KeyAscii = 0
End Sub

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
