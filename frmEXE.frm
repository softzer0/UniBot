VERSION 5.00
Begin VB.Form frmEXE 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Make EXE bot"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4455
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
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   2640
      TabIndex        =   8
      Text            =   "My bot"
      Top             =   650
      Width           =   1695
   End
   Begin VB.CheckBox chkMT 
      Caption         =   "Tra&y"
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      ToolTipText     =   "Minimized to tray on start"
      Top             =   360
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Icon"
      Height          =   1095
      Left            =   2160
      TabIndex        =   15
      Top             =   960
      Width           =   2175
      Begin VB.PictureBox picIcon 
         AutoRedraw      =   -1  'True
         ClipControls    =   0   'False
         Height          =   780
         Index           =   1
         Left            =   120
         Picture         =   "frmEXE.frx":0000
         ScaleHeight     =   505.68
         ScaleMode       =   0  'User
         ScaleWidth      =   505.68
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Width           =   780
      End
      Begin VB.PictureBox picIcon 
         AutoRedraw      =   -1  'True
         ClipControls    =   0   'False
         Height          =   780
         Index           =   0
         Left            =   120
         ScaleHeight     =   505.68
         ScaleMode       =   0  'User
         ScaleWidth      =   505.68
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   780
      End
      Begin VB.CommandButton cmdChoose 
         Caption         =   "C&hoose..."
         Height          =   255
         Left            =   1030
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.OptionButton optSilent 
      Caption         =   "&Silent"
      Height          =   255
      Left            =   2800
      TabIndex        =   5
      Top             =   360
      Width           =   735
   End
   Begin VB.OptionButton optGUI 
      Caption         =   "&GUI"
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   360
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.Frame frm1 
      Caption         =   "Include plugins:"
      Height          =   2295
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1935
      Begin VB.CommandButton cmdDetect 
         Caption         =   "&Auto-detect"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton optTemp 
         Caption         =   "T&emp"
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   1960
         Width           =   735
      End
      Begin VB.OptionButton optCurrent 
         Caption         =   "&Current"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1960
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.ListBox lstPlugins 
         Height          =   1230
         Left            =   120
         MultiSelect     =   1  'Simple
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lbl1 
         Caption         =   "Drop them into folder:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Except .NET RegEx (if it's selected) which must be dropped in directory where's EXE."
         Top             =   1725
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdProceed 
      Caption         =   "Proceed"
      Default         =   -1  'True
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label lblT 
      Caption         =   "&Title:"
      Height          =   255
      Left            =   2180
      TabIndex        =   7
      Top             =   660
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Default start-up parameter:"
      Height          =   255
      Left            =   2160
      TabIndex        =   14
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmEXE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Const FILE_SHARE_READ = &H1
Private Const GENERIC_READ              As Long = &H80000000
Private Const OPEN_EXISTING             As Long = &H3

Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Const LR_LOADFROMFILE As Long = &H10
Private Const DI_NORMAL As Long = &H3

Private Sub cmdChoose_Click()
Dim strT As String: strT = CommDlg(, "Select icon file", "ICO file (*.ico)|*.ico", , "icon")
If strT = vbNullString Then Exit Sub
Dim hIcon As Long
picIcon(0).Picture = Nothing
hIcon = LoadImageAsString(0&, strT, IMAGE_ICON, 0&, 0&, LR_LOADFROMFILE)
If hIcon Then
DrawIconEx picIcon(0).hDC, 0, 0, hIcon, 0&, 0&, 0&, 0&, DI_NORMAL
DestroyIcon hIcon
picIcon(0).Tag = strT
If picIcon(1).Visible Then picIcon(1).Visible = False Else: CloseHandle frmMain.lngIF
frmMain.lngIF = CreateFile(strT, GENERIC_READ, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, 0, ByVal 0&)
Else
picIcon(0).Tag = vbNullString
picIcon(1).Visible = True
MsgBox "Invalid icon!", vbExclamation
End If
End Sub

Private Sub cmdDetect_Click()
cmdDetect.Enabled = False
cmdProceed.Enabled = False
cmdChoose.Enabled = False
Form_KeyPress 4
frmMain.DetectP
End Sub

Private Sub cmdProceed_Click()
If cmdDetect.Enabled Then If MsgBox("You haven't checked for plugins with auto-detection option. Continue anyway?", vbExclamation + vbYesNo) = vbNo Then Exit Sub
Dim strT As String: strT = CommDlg(True, "Select where to save EXE bot", "Application (*.exe)|*.exe", , Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(txtTitle.Text, """", vbNullString), "*", vbNullString), "/", vbNullString), ":", vbNullString), ">", vbNullString), "<", vbNullString), "?", vbNullString), "\", vbNullString), "|", vbNullString))
If strT <> vbNullString Then frmMain.BuildEXE strT
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
Dim i As Integer
If KeyAscii = 1 Then
If ActiveControl = txtTitle Then
txtTitle.SelStart = 0
txtTitle.SelLength = Len(txtTitle.Text)
KeyAscii = 0
ElseIf ActiveControl = lstPlugins Then
For i = 0 To lstPlugins.ListCount - 1
lstPlugins.Selected(i) = True
Next i
End If
ElseIf KeyAscii = 4 Then
For i = 0 To lstPlugins.ListCount - 1
lstPlugins.Selected(i) = False
Next i
End If
End Sub

Private Sub Form_Load()
If frmMain.chkOnTop.Checked Then SetTopMostWindow Me.hwnd, True
LoadLst lstPlugins
If frmMain.cmdSaveC.Tag <> vbNullString Then
txtTitle.Text = Left$(frmMain.Caption, InStrRev(frmMain.Caption, "- UniBot") - 1) & "bot"
If Left$(txtTitle.Text, 1) = "*" Then txtTitle.Text = "Modified " & Mid$(txtTitle.Text, 2)
End If
'lstPlugins.Selected(0) = True 'cmdDetect_Click 'del
If frmMain.bolChk And lstPlugins.ListCount > 0 Then Exit Sub
If Not frmMain.bolDebug Or Not frmMain.bolChk Then cmdDetect.Enabled = False
If frmMain.strPl <> vbLf And lstPlugins.ListCount > 0 Then
Dim s() As String, i As Byte
s() = Split(frmMain.strPl, vbLf)
For i = 1 To UBound(s) - 1
lstPlugins.ListIndex = -1
SetP Left$(s(i), InStr(s(i), "|") - 1), Mid$(s(i), InStr(s(i), "|") + 1)
Next i
'If Not frmMain.bolChk Then lstPlugins.Enabled = False
Else
If Not frmMain.bolDebug Then frm1.Enabled = False
lstPlugins.Enabled = False
lbl1.Enabled = False
optCurrent.Enabled = False
optTemp.Enabled = False
End If
End Sub

Sub SetP(strT As String, bytI As Byte)
Dim i As Integer
For i = bytI To lstPlugins.ListCount - 1
If lstPlugins.list(i) = strT Then
lstPlugins.Selected(i) = True
Exit Sub
End If
Next
If Not frmMain.bolChk Then
cmdDetect.Enabled = True
frmMain.bolChk = True
frmMain.strPl = vbLf
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not picIcon(1).Visible And Me.Visible Then CloseHandle frmMain.lngIF
End Sub

Private Sub optGUI_Click()
If chkMT.Caption = "Tra&y" Then Exit Sub
chkMT.Caption = "Tra&y"
chkMT.ToolTipText = "Minimized to tray on start"
End Sub

Private Sub optSilent_Click()
If chkMT.Caption = "&Melt" Then Exit Sub
chkMT.Caption = "&Melt"
chkMT.ToolTipText = "Self-deletion after execution."
End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)
If KeyAscii <> 1 Then Exit Sub
KeyAscii = 0
txtTitle.SelStart = 0
txtTitle.SelLength = Len(txtTitle.Text)
End Sub
