VERSION 5.00
Begin VB.Form frmB 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Display"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   7095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton cmdOpen 
      Caption         =   "If it doesn't work, click here to open in browser."
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
      Left            =   50
      TabIndex        =   0
      Top             =   4560
      Width           =   7050
   End
End
Attribute VB_Name = "frmB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ctlD As VBControlExtender, strURL As String, bolE As Boolean, intD(1) As Integer
Public bolU As Boolean, strSrc As String

Private Sub cmdOpen_Click()
If strURL = "file:///%tmp%\UniBot.html" Then
On Error GoTo E
frmMain.PutContents Environ("tmp") & "\UniBot.html", strSrc, IIf(IsUnicode(strSrc), CP_UTF16_LE, CP_ACP)
End If
Shell "cmd.exe /c START " & strURL, vbHide
If Me.Visible Then Unload Me
Exit Sub
E:
MsgBox "There was some error in creating temporary file!", vbCritical
If Me.Visible Then Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
If frmMain.chkOnTop.Checked Then SetTopMostWindow Me.hWnd, True
If bolU Then Exit Sub
If strSrc <> vbNullString Then strURL = "file:///%tmp%\UniBot.html" Else: If Left$(frmWizard.lstFN(1).list(frmWizard.lstFN(1).ListIndex), 8) <> "https://" And Left$(frmWizard.lstFN(1).list(frmWizard.lstFN(1).ListIndex), 7) <> "http://" Then strURL = frmWizard.txtSite.Text & frmWizard.lstFN(1).list(frmWizard.lstFN(1).ListIndex) Else: strURL = frmWizard.lstFN(1).list(frmWizard.lstFN(1).ListIndex)
On Error GoTo E
Set ctlD = Controls.add("Shell.Explorer.2", "w")
ctlD.Move 50, 50, Me.Width, Me.Height - 300
ctlD.Visible = True
ctlD.object.Silent = True
If strSrc <> vbNullString Then
ctlD.object.Navigate "about:blank"
ctlD.object.Document.Open
ctlD.object.Document.Write strSrc
ctlD.object.Document.Close
Else: ctlD.object.Navigate strURL
End If
If intD(0) > 0 Then
Me.Width = intD(0)
Me.Height = intD(1)
End If
Me.Show vbModal
Exit Sub
E:
If Not bolE Then MsgBox "There is some unexpected error! It will be opened in browser.", vbExclamation: bolE = True
cmdOpen_Click
bolU = True
End Sub

Private Sub Form_Resize()
On Error Resume Next
ctlD.Width = Me.ScaleWidth - 100
ctlD.Height = Me.ScaleHeight - 350
cmdOpen.Top = Me.ScaleHeight - 300
cmdOpen.Width = Me.ScaleWidth - 95
End Sub

Private Sub Form_Unload(Cancel As Integer)
intD(0) = Me.Width
intD(1) = Me.Height
strSrc = vbNullString
End Sub
