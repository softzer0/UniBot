VERSION 5.00
Begin VB.Form frmT 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fine tuning"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3810
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
   ScaleHeight     =   3450
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Trimming && other"
      Height          =   1545
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   3575
      Begin VB.CheckBox chkDebug 
         Caption         =   "&Debug mode"
         Height          =   255
         Left            =   1990
         TabIndex        =   15
         Top             =   860
         Width           =   1215
      End
      Begin VB.CheckBox chkColl 
         Caption         =   "Don't avoid &collisions"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   860
         Width           =   1935
      End
      Begin VB.TextBox txtOutMax 
         Height          =   285
         Left            =   2805
         TabIndex        =   19
         ToolTipText     =   "If output data is important then leave zero."
         Top             =   1160
         Width           =   615
      End
      Begin VB.TextBox txtLogMax 
         Height          =   285
         Left            =   980
         TabIndex        =   17
         ToolTipText     =   "Maximum 32767 (zero can't be left)."
         Top             =   1160
         Width           =   615
      End
      Begin VB.CheckBox chkEach 
         Caption         =   "On &each thread"
         Height          =   255
         Left            =   1990
         TabIndex        =   13
         ToolTipText     =   "If this is unchecked then number will be divided with number of threads."
         Top             =   560
         Width           =   1455
      End
      Begin VB.TextBox txtTOrigin 
         Height          =   285
         Index           =   1
         Left            =   1510
         TabIndex        =   12
         ToolTipText     =   "...nor very low or very high values are here."
         Top             =   560
         Width           =   375
      End
      Begin VB.TextBox txtTOrigin 
         Height          =   285
         Index           =   0
         Left            =   3050
         TabIndex        =   10
         ToolTipText     =   "Zero (unlimited) is not recommended for origin..."
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "&Output count:"
         Height          =   255
         Left            =   1710
         TabIndex        =   18
         ToolTipText     =   "Line count"
         Top             =   1160
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "&Log count:"
         Height          =   255
         Left            =   150
         TabIndex        =   16
         Top             =   1160
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "For &publ. str, src:"
         Height          =   255
         Left            =   150
         TabIndex        =   11
         ToolTipText     =   "For public strings, and source"
         Top             =   560
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "&Maximum origin depth level for display:"
         Height          =   260
         Left            =   156
         TabIndex        =   9
         Top             =   234
         Width           =   2925
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   3090
      Width           =   3570
   End
   Begin VB.Frame Frame1 
      Caption         =   "Output template"
      Height          =   910
      Left            =   120
      TabIndex        =   3
      Top             =   450
      Width           =   3575
      Begin VB.TextBox txtTemplate 
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   7
         Top             =   530
         Width           =   2735
      End
      Begin VB.TextBox txtTemplate 
         Height          =   285
         Index           =   0
         Left            =   720
         TabIndex        =   5
         ToolTipText     =   "{T} - thread, {S} - sub-thread, {I} - index, {O} - origin, {D} - current date & time, {N} - current string name, [nl] - new line"
         Top             =   240
         Width           =   2735
      End
      Begin VB.Label Label3 
         Caption         =   "&After:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   530
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "&Before:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.ComboBox cmbHS 
      Height          =   315
      ItemData        =   "frmTuning.frx":0000
      Left            =   2840
      List            =   "frmTuning.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtAfter 
      Height          =   285
      Left            =   2240
      TabIndex        =   1
      ToolTipText     =   "Leave zero for infinite time."
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "&Stop execution of bot after:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intAfter As Integer
Public bolHours As Boolean
Public strTemplate0 As String, strTemplate1 As String
Public bytTOrigin0 As Byte, bytTOrigin1 As Byte
Public bolNoEach As Boolean
Public intLogMax As Integer
Public intOutMax As Integer
Public bolColl As Boolean

Private Sub cmdOK_Click()
If txtAfter.Text <= 32767 Then intAfter = txtAfter.Text Else: intAfter = 32767
bolHours = cmbHS.ListIndex = 1
strTemplate0 = Replace(txtTemplate(0).Text, "[nl]", vbNewLine)
strTemplate1 = Replace(txtTemplate(1).Text, "[nl]", vbNewLine)
If txtTOrigin(0).Text <= 255 Then bytTOrigin0 = txtTOrigin(0).Text Else: bytTOrigin0 = 255
If txtTOrigin(1).Text <= 255 Then bytTOrigin1 = txtTOrigin(1).Text Else: bytTOrigin1 = 255
bolNoEach = Not CBool(chkEach.Value)
bolColl = CBool(chkColl.Value)
If txtLogMax.Text > 0 And txtLogMax.Text <= 32767 Then
intLogMax = txtLogMax.Text
If intLogMax < 32767 Then
frmMain.lblStatus.Caption = "Trimming log..."
frmMain.lblStatus.Refresh
Do While frmMain.lstLog.ListCount > intLogMax
frmMain.lstLog.RemoveItem 0
Loop
End If
Else: intLogMax = 32767
End If
If txtOutMax.Text <= 32767 Then
intOutMax = txtOutMax.Text
If intOutMax > 0 Then
frmMain.lblStatus.Caption = "Trimming output..."
frmMain.lblStatus.Refresh
Do While (Len(frmMain.txtOutput.Text) - Len(Replace(frmMain.txtOutput.Text, vbNewLine, vbNullString))) / 2 > intOutMax
frmMain.txtOutput.Text = Mid$(frmMain.txtOutput.Text, InStr(frmMain.txtOutput.Text, vbNewLine) + 2)
Loop
End If
Else: intOutMax = 32767
End If
frmMain.bolDebug = CBool(chkDebug.Value)
frmMain.lblStatus.Caption = "Idle..."
Unload Me
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
            ActiveControl.SelStart = 0
            ActiveControl.SelLength = Len(ActiveControl.Text)
        End If
    End If
End Sub

Private Sub Form_Load()
If frmMain.chkOnTop.Checked Then SetTopMostWindow Me.hWnd, True
txtAfter.Text = intAfter
If bolHours Then cmbHS.ListIndex = 1 Else: cmbHS.ListIndex = 0
txtTemplate(0).Text = Replace(strTemplate0, vbNewLine, "[nl]")
txtTemplate(1).Text = Replace(strTemplate1, vbNewLine, "[nl]")
txtTOrigin(0).Text = bytTOrigin0
txtTOrigin(1).Text = bytTOrigin1
chkEach.Value = CInt(Not bolNoEach) * (-1)
chkColl.Value = CInt(bolColl) * (-1)
chkColl.Enabled = Not txtTOrigin(1).Text = "0"
txtLogMax.Text = intLogMax
txtOutMax.Text = intOutMax
chkDebug.Value = CInt(frmMain.bolDebug) * (-1)
End Sub

Private Sub txtAfter_Change()
CheckText txtAfter
End Sub

Private Sub txtLogMax_Change()
CheckText txtLogMax
End Sub

Private Sub txtTOrigin_Change(Index As Integer)
CheckText txtTOrigin(Index)
If Index = 0 Then Exit Sub
Select Case txtTOrigin(1).Text
Case "0"
If chkColl.Enabled Then
If chkColl.Value = 0 Then chkColl.Value = 1 Else: chkColl.Tag = "1"
chkColl.Enabled = False
End If
Case Else
If Not chkColl.Enabled Then
chkColl.Enabled = True
If chkColl.Tag = "1" Then chkColl.Tag = vbNullString Else: chkColl.Value = 0
End If
End Select
End Sub

Private Sub txtOutMax_Change()
CheckText txtOutMax
End Sub
