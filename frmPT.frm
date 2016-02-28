VERSION 5.00
Begin VB.Form frmPT 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Proxy & thread settings"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2895
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
   ScaleHeight     =   4215
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   945
      Left            =   120
      TabIndex        =   21
      Top             =   2120
      Width           =   2710
      Begin VB.CheckBox chkChangeP 
         Caption         =   "&Chg. pr. w/ max:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   80
         TabIndex        =   10
         ToolTipText     =   "Change proxy with maximum:"
         Top             =   600
         Width           =   1545
      End
      Begin VB.TextBox txtCycles 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1640
         TabIndex        =   12
         ToolTipText     =   "Leave zero to change on every request."
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtMaxR 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2105
         TabIndex        =   9
         ToolTipText     =   "How many retries in cycle per request? Leave zero for unlimited."
         Top             =   260
         Width           =   495
      End
      Begin VB.CheckBox chkRetry 
         Caption         =   "On error &retry:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   0
         Width           =   1455
      End
      Begin VB.TextBox txtDelay 
         Enabled         =   0   'False
         Height          =   285
         Left            =   675
         TabIndex        =   7
         ToolTipText     =   "In seconds. Leave zero to disable (not recommended)."
         Top             =   260
         Width           =   495
      End
      Begin VB.Label lblY 
         Caption         =   "c&ycles"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2190
         TabIndex        =   11
         Top             =   620
         Width           =   495
      End
      Begin VB.Label lblD 
         Caption         =   "&Delay:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   275
         Width           =   495
      End
      Begin VB.Label lblM 
         Caption         =   "&Maximum:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1300
         TabIndex        =   8
         Top             =   275
         Width           =   735
      End
   End
   Begin VB.CheckBox chkStartP 
      Caption         =   "S&tart w/ pr."
      Height          =   255
      Left            =   1560
      TabIndex        =   13
      ToolTipText     =   "Start with proxy"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtTimeout 
      Height          =   285
      Left            =   840
      TabIndex        =   15
      ToolTipText     =   "second(s)"
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox txtThreads 
      Height          =   285
      Left            =   840
      TabIndex        =   17
      ToolTipText     =   "Maximum 254 (together with sub-threads)."
      Top             =   3480
      Width           =   495
   End
   Begin VB.CheckBox chkSame 
      Caption         =   "S&ame for each thread"
      Height          =   220
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   1900
   End
   Begin VB.CheckBox chkSkip 
      Caption         =   "Sk&ip bad pr."
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      ToolTipText     =   "Skip bad proxies"
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtSubThr 
      Height          =   285
      Left            =   2280
      TabIndex        =   19
      ToolTipText     =   "Can cause instability if zero is left (for unlimited)."
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "O&K"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3840
      Width           =   2655
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtProxies 
      Height          =   1305
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   390
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "Tim&eout:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "&Sub-thr.:"
      Height          =   255
      Left            =   1560
      TabIndex        =   18
      ToolTipText     =   "Maximum sub-threads (if needed):"
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "T&hreads:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "&Proxies:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strProxy As String
Public bolSame As Boolean
Public bolSkip As Boolean
Public bolNoStartP As Boolean
Public bolNoRetry As Boolean
Public bolNoChange As Boolean
Public bytTimeout As Byte
Public bytThreads As Byte
Public bytSubThr As Byte
Public bytDelay As Byte
Public bytMaxR As Byte
Public bytCycles As Byte

Private Sub chkChangeP_Click()
If chkRetry.Value = 0 Then Exit Sub
If chkChangeP.Value = 1 Then
lblY.Enabled = True
txtCycles.Enabled = True
Else
lblY.Enabled = False
txtCycles.Enabled = False
End If
End Sub

Private Sub chkRetry_Click()
If chkRetry.Value = 1 Then
If chkChangeP.Value = 1 Then
lblY.Enabled = True
txtCycles.Enabled = True
End If
chkChangeP.Enabled = True
lblD.Enabled = True
txtDelay.Enabled = True
lblM.Enabled = True
txtMaxR.Enabled = True
Else
If chkChangeP.Value = 1 Then
lblY.Enabled = False
txtCycles.Enabled = False
End If
chkChangeP.Enabled = False
lblD.Enabled = False
txtDelay.Enabled = False
lblM.Enabled = False
txtMaxR.Enabled = False
End If
End Sub

Private Sub cmdOK_Click()
Dim lngT As Long
If txtThreads.Text < 255 Then
lngT = bytThreads <> txtThreads.Text
bytThreads = txtThreads.Text
Else
lngT = bytThreads <> 254
bytThreads = 254
End If
If txtProxies.Text <> strProxy Then
'Me.Enabled = False
'Dim bolP As Boolean
'If frmMain.bolDebug And strProxy <> vbNullString Then bolP = True
frmMain.lblStatus.Caption = "Processing..."
frmMain.lblStatus.Refresh
Screen.MousePointer = 11
Dim strT As String: strT = Replace(Replace(txtProxies.Text, vbCr, vbNullString), vbLf, vbNewLine)
Do While InStr(strT, vbNewLine & vbNewLine) > 0
strT = Replace(strT, vbNewLine & vbNewLine, vbNewLine)
Loop
Do While Right$(strT, 2) = vbNewLine
strT = Left$(strT, Len(strT) - 2)
Loop
Do While Left$(strT, 2) = vbNewLine
strT = Mid$(strT, 3)
Loop
lngT = UBound(Split(Replace(strT, vbCr, vbNullString), vbLf)) + 1
If lngT > 0 Then GoSub ChkN
If frmMain.bolDebug Then
If strT <> vbNullString Then
If strT <> strProxy Then frmMain.addLog "Potential proxies set.", True
ElseIf strProxy <> vbNullString Then frmMain.addLog "Proxy list cleared.", True
End If
End If
strProxy = strT 'strProxy = vbNullString
'With frmMain.lblStatus
'.Caption = "Checking validity of proxy addresses"
'.Caption = .Caption & " and removing duplicates" 'check if option to remove duplicates is enabled
'.Caption = .Caption & "..."
'End With
'frmMain.lblStatus.Caption = "Removing duplicates and blank lines..."
'Screen.MousePointer = 11
'DoEvents
'Dim s() As String: s() = Split(Replace(txtProxies.Text, vbCr, vbNullString), vbLf)
'For i = 0 To UBound(s())
's(i) = Trim$(s(i))
'If s(i) <> vbNullString Then If InStr(strProxy, s(i)) = 0 Then strProxy = strProxy & s(i) & vbNewLine 'If RegExpr("\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b:\d{2,5}", s(i), , 0) Then
'Next
'If strProxy <> vbNullString Then
'strProxy = Left$(strProxy, Len(strProxy) - 2)
'If frmMain.bolDebug Then If UBound(Split(strProxy, vbNewLine)) = 0 Then frmMain.addLog "1 proxy is ready for use.", True Else: frmMain.addLog UBound(Split(strProxy, vbNewLine)) + 1 & " proxies are ready for use.", True
'Else
'If frmMain.bolDebug Then
'^ 'If bolP Then
'End If
'End If
'Me.Enabled = True
'Screen.MousePointer = 0
'frmMain.lblStatus.Caption = "Idle..."
ElseIf lngT = -1 Then
lngT = UBound(Split(Replace(strProxy, vbCr, vbNullString), vbLf)) + 1
If lngT > 0 Then GoSub ChkN
End If
bolSame = CBool(chkSame.Value)
bolSkip = CBool(chkSkip.Value)
bolNoStartP = Not CBool(chkStartP.Value)
bolNoRetry = Not CBool(chkRetry.Value)
bolNoChange = Not CBool(chkChangeP.Value)
If txtSubThr.Text < 255 Then bytSubThr = txtSubThr.Text Else: bytSubThr = 254
If CInt(bytThreads) + CInt(bytSubThr) > 255 Then If bytSubThr > bytThreads Then bytSubThr = bytSubThr - bytThreads Else: bytSubThr = 1
If txtTimeout.Text < 256 Then bytTimeout = txtTimeout.Text Else: bytTimeout = 255
If txtDelay.Text < 256 Then bytDelay = txtDelay.Text Else: bytDelay = 255
If txtMaxR.Text < 256 Then bytMaxR = txtMaxR.Text Else: bytMaxR = 255
If txtCycles.Text < 256 Then bytCycles = txtCycles.Text Else: bytCycles = 255
frmMain.lblStatus.Caption = "Idle..."
Screen.MousePointer = 0
Unload Me
Exit Sub
ChkN:
If Not bolSame Then
If bytThreads > lngT Then
Screen.MousePointer = 0
Select Case MsgBox("Insufficient amount of potential proxies! Yes to continue anyway, No to decrease threads to their count.", vbExclamation + vbYesNoCancel)
Case vbCancel
frmMain.lblStatus.Caption = "Idle..."
Exit Sub
Case vbNo: bytThreads = lngT
Case vbYes: Screen.MousePointer = 11
End Select
End If
End If
Return
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

Private Sub cmdLoad_Click()
Dim strFile As String
strFile = CommDlg(, "Select proxy list file to load", "Text file (*.txt)|*.txt|Any file|*.*", , "proxies")
If strFile = vbNullString Then Exit Sub
frmMain.lblStatus.Caption = "Loading file for proxy list..."
Screen.MousePointer = 11
frmMain.lblStatus.Refresh
Dim strT As String
If Trim$(Replace(Replace(txtProxies.Text, vbCr, vbNullString), vbLf, vbNullString)) <> vbNullString Then
If InStr(txtProxies.Text, vbLf) > 0 Then
If Split(txtProxies.Text, vbLf)(UBound(Split(txtProxies.Text, vbLf))) <> vbNullString Then GoTo C
Else
C: strT = vbNewLine
End If
Else: txtProxies.Text = vbNullString
End If
LoadBig txtProxies, strFile, , txtProxies.Text & strT
If frmMain.bolDebug Then frmMain.addLog "File """ & frmMain.get_relative_path_to(strFile) & """ loaded for proxy list.", True
Screen.MousePointer = 0
frmMain.lblStatus.Caption = "Idle..."
End Sub

Private Sub Form_Load()
If frmMain.chkOnTop.Checked Then SetTopMostWindow Me.hWnd, True
LoadBig txtProxies, strProxy, True
chkSame.Value = CInt(bolSame) * (-1)
chkSkip.Value = CInt(bolSkip) * (-1)
chkStartP.Value = CInt(Not bolNoStartP) * (-1)
chkRetry.Value = CInt(Not bolNoRetry) * (-1)
chkChangeP.Value = CInt(Not bolNoChange) * (-1)
txtCycles.Text = bytCycles
txtDelay.Text = bytDelay
txtMaxR.Text = bytMaxR
txtTimeout.Text = bytTimeout
txtThreads.Text = bytThreads
txtSubThr.Text = bytSubThr
End Sub

Private Sub txtCycles_Change()
CheckText txtCycles
End Sub

Private Sub txtDelay_Change()
CheckText txtDelay
End Sub

Private Sub txtMaxR_Change()
CheckText txtMaxR
End Sub

Private Sub txtTimeout_Change()
CheckText txtTimeout
End Sub

Private Sub txtThreads_Change()
CheckText txtThreads
End Sub

Private Sub txtSubThr_Change()
CheckText txtSubThr
End Sub

Private Sub txtProxies_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 22 Or KeyAscii = 3 Or KeyAscii = 19 Or KeyAscii = 24 Then Exit Sub
If InStr("0123456789:." & vbNewLine, Chr$(KeyAscii)) > 0 Then Exit Sub
KeyAscii = 0
End Sub
