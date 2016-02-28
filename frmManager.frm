VERSION 5.00
Begin VB.Form frmManager 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Index manager"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   1260
      Width           =   735
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   900
      Width           =   735
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "&End"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "&Down"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "&Up"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.ListBox lstI 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bolL As Boolean
Dim bytL As Byte

Private Sub cmdAdd_Click()
Dim p As Byte
If lstI.ListIndex = lstI.ListCount - 1 + bytL Then frmMain.cmbIndex.ListIndex = lstI.ListCount - 2 Else: frmMain.cmbIndex.ListIndex = lstI.ListIndex: If lstI.ListIndex = lstI.ListCount - 2 Then Me.Tag = "-"
bolL = True
frmMain.cmdI_Click
bolL = False
If Me.Tag = "  " Then
Me.Tag = vbNullString
bytL = 0
GoTo E
ElseIf Me.Tag = " " Then
Me.Tag = vbNullString
If cmdAdd.Caption <> "&Add" Then If frmMain.Filled(lstI.ListCount - 1) Then bytL = 1: lstI_Click
lstI.SetFocus
Exit Sub
End If
If lstI.ListIndex = lstI.ListCount - 1 + bytL Then
lstI.AddItem lstI.ListCount & Mid$(lstI.list(lstI.ListCount - 1), InStr(lstI.list(lstI.ListCount - 1), vbTab))
lstI.RemoveItem lstI.ListCount - 2
If frmMain.Filled(lstI.ListCount - 1) Then bytL = 1 Else: bytL = 0
E:
lstI.AddItem lstI.ListCount + 1 & vbTab
lstI.ListIndex = lstI.ListCount - 1
lstI.SetFocus
Else: Unload Me
End If
End Sub

Private Sub cmdRemove_Click()
frmMain.cmbIndex.ListIndex = lstI.ListIndex
bolL = True
frmMain.cmdR_Click
bolL = False
If StrPtr(Me.Tag) = 0 Then lstI.SetFocus: Exit Sub
If frmMain.cmbIndex.ListCount = 1 Then Unload Me: Exit Sub
Me.Tag = vbNullString
bytL = 0
Dim p As Byte, a As Byte
p = lstI.ListIndex
lstI.RemoveItem p
SetListboxScrollbar1 lstI
For a = p To lstI.ListCount - 1
ChO a
Next
If frmMain.Filled(lstI.ListCount - 1) Then lstI.AddItem lstI.ListCount + 1 & vbTab
lstI.ListIndex = p
lstI_Click
lstI.SetFocus
End Sub

Private Sub cmdStart_Click()
StartEnd
End Sub

Private Sub cmdUp_Click()
UpDown
End Sub

Private Sub cmdDown_Click()
UpDown -1
End Sub

Private Sub cmdEnd_Click()
StartEnd -1
End Sub

Private Sub StartEnd(Optional bytT As Integer = 1)
frmMain.lblStatus.Caption = "Shifting indexes..."
frmMain.lblStatus.Refresh
Screen.MousePointer = 11
Dim p As Byte, s As Byte
p = lstI.ListIndex
If bytT = -1 Then s = lstI.ListCount - 2 + bytL Else: s = 0
lstI.Tag = Mid$(lstI.list(p), InStr(lstI.list(p), vbTab) + 1)
frmMain.SLInd p
lstI.RemoveItem p
lstI.AddItem CInt(s) + 1 & vbTab & lstI.Tag, s
If frmMain.cmbIndex.ListIndex = p Then frmMain.cmbIndex.ListIndex = p - bytT: frmMain.cmbIndex.Tag = p Else: If frmMain.cmbIndex.ListIndex = s Then frmMain.cmbIndex.ListIndex = s + bytT: frmMain.cmbIndex.Tag = "-1"
Do While p <> s
frmMain.ChngI p, p - bytT
ChO p
p = p - bytT
Loop
frmMain.SLInd s, True
If frmMain.cmbIndex.Tag <> vbNullString Then If frmMain.cmbIndex.Tag <> "-1" Then frmMain.cmbIndex.ListIndex = frmMain.cmbIndex.Tag: frmMain.cmbIndex.Tag = vbNullString Else: frmMain.cmbIndex.ListIndex = s: frmMain.cmbIndex.Tag = vbNullString
lstI.ListIndex = s
Cmp
End Sub

Private Sub UpDown(Optional bytT As Integer = 1)
frmMain.lblStatus.Caption = "Shifting indexes..."
frmMain.lblStatus.Refresh
Screen.MousePointer = 11
Dim p As Byte: p = lstI.ListIndex
lstI.Tag = Mid$(lstI.list(p), InStr(lstI.list(p), vbTab) + 1)
frmMain.SLInd p
lstI.RemoveItem p
lstI.AddItem (CInt(p) - bytT) + 1 & vbTab & lstI.Tag, p - bytT
frmMain.ChngI p, p - bytT
If frmMain.cmbIndex.ListIndex = p Then frmMain.cmbIndex.ListIndex = p - bytT: frmMain.cmbIndex.Tag = "0" Else: If frmMain.cmbIndex.ListIndex = p - bytT Then frmMain.cmbIndex.ListIndex = p: frmMain.cmbIndex.Tag = "1"
frmMain.SLInd p - bytT, True
If frmMain.cmbIndex.Tag = "0" Then frmMain.cmbIndex.ListIndex = p: frmMain.cmbIndex.Tag = vbNullString Else: If frmMain.cmbIndex.Tag = "1" Then frmMain.cmbIndex.ListIndex = p - bytT: frmMain.cmbIndex.Tag = vbNullString
ChO p
lstI.ListIndex = p - bytT
Cmp
End Sub

Private Sub Cmp()
If Left$(frmMain.Caption, 1) <> "*" Then frmMain.Caption = "*" & frmMain.Caption
Screen.MousePointer = 0
frmMain.lblStatus.Caption = "Idle..."
lstI.SetFocus
End Sub

Private Sub ChO(p As Byte)
lstI.Tag = Mid$(lstI.list(p), InStr(lstI.list(p), vbTab))
lstI.RemoveItem p
lstI.AddItem CInt(p) + 1 & lstI.Tag, p
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
If frmMain.chkOnTop.Checked Then SetTopMostWindow Me.hWnd, True
bytL = 0
If frmMain.bytLimit = frmMain.cmbIndex.ListCount - 1 Then If frmMain.Filled(frmMain.cmbIndex.ListCount - 1) Then bytL = 1
End Sub

Private Sub lstI_Click()
If lstI.ListIndex = 0 Or lstI.ListIndex = lstI.ListCount - 1 + bytL Then cmdUp.Enabled = False Else: cmdUp.Enabled = True
If lstI.ListIndex < 2 Or lstI.ListIndex = lstI.ListCount - 1 + bytL Then cmdStart.Enabled = False Else: cmdStart.Enabled = True
If lstI.ListIndex > lstI.ListCount - 4 + bytL Then cmdEnd.Enabled = False Else: cmdEnd.Enabled = True
If lstI.ListIndex > lstI.ListCount - 3 + bytL Then cmdDown.Enabled = False Else: cmdDown.Enabled = True
If lstI.ListIndex = lstI.ListCount - 1 + bytL Then cmdRemove.Enabled = False: cmdAdd.Caption = "&Add pr.": cmdAdd.ToolTipText = "Duplicate previous index" Else: cmdRemove.Enabled = True: cmdAdd.Caption = "&Add": cmdAdd.ToolTipText = vbNullString
End Sub

Private Sub lstI_DblClick()
frmMain.cmbIndex.ListIndex = lstI.ListIndex
Unload Me
End Sub

Private Sub lstI_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And lstI.ListIndex <> -1 Then lstI_DblClick
End Sub
