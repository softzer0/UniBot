VERSION 5.00
Begin VB.Form frmEE 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4320
   ClientLeft      =   0
   ClientTop       =   -60
   ClientWidth     =   3855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   288
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   257
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   840
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   3240
      Width           =   3375
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Run it in Windows for full functionality :)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   3135
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   3120
      Top             =   2640
   End
End
Attribute VB_Name = "frmEE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long

Private C As Cube

Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal crKey As Long, _
    ByVal bAlpha As Byte, _
    ByVal dwFlags As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1&
Private Const LWA_ALPHA = &H2&
Private Const LWA_BOTH = 3
Dim BackColor1 As Long
Dim trans As Byte

Dim s(1) As Long, s1(1) As Long, s2(1) As Boolean
Attribute s.VB_VarHelpID = -1
Public unl As Boolean

Private Sub SetTrans(hWnd As Long, trans As Byte)
    Dim Tcall As Long
    Tcall = GetWindowLong(hWnd, GWL_EXSTYLE)
    SetWindowLong hWnd, GWL_EXSTYLE, Tcall Or WS_EX_LAYERED
    SetLayeredWindowAttributes hWnd, BackColor1, trans, LWA_ALPHA + LWA_COLORKEY
End Sub

Private Sub Form_Activate()
While IsWindowVisible(hWnd) And Not unl
    C.Roll = C.Roll + 0.1
    C.Pitch = C.Pitch + 0.1
    C.Yaw = C.Yaw + 0.1

    DoEvents
   
    C.Draw Me.hDC
Wend
End Sub

Private Sub Form_Load()
BackColor1 = RGB(127, 127, 0)
With Me
.Top = (Screen.Height - .Height) / 2
.Left = (Screen.Width - .Width) / 2
End With
SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
SetTrans Me.hWnd, 0
Me.BackColor = BackColor1
Frame1.BackColor = BackColor1
Label1.BackColor = BackColor1
Timer1.Enabled = True

AddW

Randomize
Set C = New Cube
C.x = ScaleWidth / 2
C.y = (ScaleHeight - 65) / 2
uFMOD_PlaySong 101, 0, XM_RESOURCE
s1(0) = Me.Top
s1(1) = Me.Left
s(0) = 25
s(1) = 21
Timer2.Enabled = True
End Sub

Function AddW()
On Error GoTo E
Dim ctlD As VBControlExtender
Set ctlD = Controls.add("Shell.Explorer.2", "w", Frame1)
ctlD.Move -120, -120, 3855, 1455
ctlD.Visible = True
ctlD.object.Silent = True
ctlD.object.Navigate "about:blank"
ctlD.object.Document.Open
ctlD.object.Document.Write StrConv(LoadResData(101, "CUSTOM"), vbUnicode)
ctlD.object.Document.Close
E:
End Function

Private Sub Timer1_Timer()
On Error GoTo E
If unl Then trans = trans - 5 Else: trans = trans + 1
SetTrans Me.hWnd, trans
Exit Sub
E:
If unl Then
uFMOD_PlaySong 0, 0, 0
Timer2.Enabled = False
unl = False
Unload Me
Else: Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
If Not s2(0) Then
Me.Top = Me.Top - 5
If Me.Top - 5 < s1(0) - s(0) Then s2(0) = True
Else
Me.Top = Me.Top + 5
If Me.Top + 5 > s1(0) + s(0) Then s2(0) = False
End If
If Not s2(1) Then
Me.Left = Me.Left - 3
If Me.Left - 3 < s1(1) - s(1) Then s2(1) = True
Else
Me.Left = Me.Left + 3
If Me.Left + 3 > s1(1) + s(1) Then s2(1) = False
End If
End Sub
