VERSION 5.00
Begin VB.Form frmD 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   1995
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   1995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkOutput 
      Caption         =   "&Output"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   810
   End
   Begin VB.CheckBox chkPublic 
      Caption         =   "&Public"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.CheckBox chkCrucial 
      Caption         =   "&Crucial"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.CheckBox chkArray 
      Caption         =   "&Array"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bolA As Boolean

Private Sub chkOutput_Click()
If Not Me.Visible Or bolA Then bolA = False: Exit Sub
If ProcS(chkOutput) = 1 Then Unload Me
End Sub

Private Sub chkPublic_Click()
If Not Me.Visible Or bolA Then bolA = False: Exit Sub
If chkPublic.Value = 0 Or Not CheckPublic(frmMain.txtString(Me.Tag).Text, CBool(chkPublic.Value)) Then
If ProcS(chkPublic) = 1 Then Unload Me
Exit Sub
End If
bolA = True
chkPublic.Value = 0
End Sub

Private Function ProcS(objC As CheckBox) As Byte
If objC.Name = "chkPublic" Then
If chkOutput.Value = 1 Then Exit Function
ElseIf chkPublic.Value = 1 Then Exit Function
End If
Select Case StrAR(frmMain.cmbIndex.ListIndex, Not CBool(objC.Value))
Case 1
bolA = True
objC.Value = 1
ProcS = 2
Case 2
bolA = True
ProcS = 1
End Select
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
If frmMain.chkOnTop.Checked Then SetTopMostWindow Me.hWnd, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not bolA Then
If chkCrucial.Value = 0 And chkPublic.Value = 0 And chkArray.Value = 0 And chkOutput.Value = 0 Then frmMain.txtExp(Me.Tag).Tag = vbNullString: Exit Sub
frmMain.txtExp(Me.Tag).Tag = chkCrucial.Value & "," & chkPublic.Value & "," & chkArray.Value & "," & chkOutput.Value
Else
frmMain.txtExp(Me.Tag).Tag = "-" & frmMain.txtExp(Me.Tag).Tag
bolA = False
End If
End Sub
