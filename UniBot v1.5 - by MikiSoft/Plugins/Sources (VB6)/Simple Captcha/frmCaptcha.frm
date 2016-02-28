VERSION 5.00
Begin VB.Form frmCaptcha 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Enter captcha"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1650
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   1650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrT 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1320
      Top             =   240
   End
   Begin VB.Timer tmrWait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1080
      Top             =   0
   End
   Begin VB.TextBox txtCaptcha 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Note: Hit Esc to stop."
      Top             =   480
      Width           =   1575
   End
   Begin VB.PictureBox picCaptcha 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   1155
      TabIndex        =   2
      Top             =   0
      Width           =   1155
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "Confirm"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblWait 
      Alignment       =   2  'Center
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
      Left            =   0
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
End
Attribute VB_Name = "frmCaptcha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2011 VBnet/Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Declare Function GetSystemMetrics Lib "user32" _
   (ByVal nIndex As Long) As Long

Private Declare Function GetMenu Lib "user32" _
   (ByVal hwnd As Long) As Long

Private Const SM_CYCAPTION = 4
Private Const SM_CYMENU = 15
Private Const SM_CXFRAME = 32

Private twipsx As Long
Private twipsy As Long

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" ( _
ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA As Long = 48
Private Type RECT
lLeft As Long
lTop As Long
lRight As Long
lBottom As Long
End Type

Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Const HWND_TOPMOST = -1

Private Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
    
Dim unl(1) As Boolean, bytR As Byte, intP As Integer
Public IsGIF As Boolean, bytC As Byte, Tess As Boolean, strI As String, strR As String, T As Byte, bytM As Byte, bytM1 As Byte
    
Private Sub Form_Activate()
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
End Sub

Private Sub Form_Load()
bytR = 0
unl(0) = False
twipsx = Screen.TwipsPerPixelX
twipsy = Screen.TwipsPerPixelY
On Error GoTo E
Dim strF As String: strF = "tmp" & App.ThreadID & Mid$(strI, 2) & ".img"
If Not IsGIF Then
Dim token As Long: token = InitGDIPlus
picCaptcha.Picture = LoadPictureGDIPlus(strF)
Call FreeGDIPlus(token)
Else: picCaptcha.Picture = LoadPicture(strF)
End If
Call AutoSizeToPicture(picCaptcha)
Me.Height = Me.Height + txtCaptcha.Height + cmdConfirm.Height + 90
txtCaptcha.Move 0, picCaptcha.Height + 20, Me.ScaleWidth
cmdConfirm.Move 0, txtCaptcha.top + txtCaptcha.Height, Me.ScaleWidth
If T > 0 Then
Me.Height = Me.Height + lblWait.Height
lblWait.Caption = T & " sec(s)"
lblWait.Move 0, cmdConfirm.top + cmdConfirm.Height + 40, Me.ScaleWidth
tmrWait.Enabled = True
Else: lblWait.Visible = False
End If
Dim deskRECT As RECT
Call SystemParametersInfo(SPI_GETWORKAREA, 0&, deskRECT, 0&)
With deskRECT
Me.Move (.lRight * Screen.TwipsPerPixelX) - Me.Width, _
(.lBottom * Screen.TwipsPerPixelY) - Me.Height
End With
If bolHid Then SetAttr strF, vbNormal
Kill strF
If Tess Then
intP = 0
tmrT.Enabled = True
End If
txtCaptcha.MaxLength = bytM1
If bytM > 0 Then cmdConfirm.Enabled = False
Exit Sub
E:
On Error GoTo -1
On Error Resume Next
If Not IsGIF Then Call FreeGDIPlus(token)
If bolHid Then SetAttr strF, vbNormal
Kill strF
CaptchaText.Add vbNullChar, strI
Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then CaptchaText.Add vbNullString, strI
End Sub

Private Sub txtCaptcha_Change()
If txtCaptcha.Text = vbNullString Then Exit Sub
txtCaptcha.Text = ProcStr(txtCaptcha.Text, , , , bytC, strR)
txtCaptcha.SelStart = Len(txtCaptcha.Text)
If Len(txtCaptcha.Text) >= bytM Then cmdConfirm.Enabled = True Else: cmdConfirm.Enabled = False
End Sub

Private Sub cmdConfirm_Click()
CaptchaText.Add txtCaptcha.Text, strI
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
Case 27
CaptchaText.Add vbNullString, strI
Unload Me
Case 1
KeyAscii = 0
txtCaptcha.SelStart = 0
txtCaptcha.SelLength = Len(txtCaptcha.Text)
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
unl(0) = True
If tmrT.Enabled Then tmrT_Timer
End Sub

Private Sub tmrT_Timer()
If unl(0) And Not unl(1) Then Exit Sub
On Error GoTo E
If tmrT.Interval = 200 Or unl(0) And unl(1) Then
If Not unl(0) Then
Open tmrT.Tag & "txt" For Input Access Read As #1
Dim tmp As String
On Error Resume Next
If Not EOF(1) Then Line Input #1, tmp
Close #1
End If
Kill tmrT.Tag & "txt"
unl(1) = False
If Not unl(0) Then
tmp = Replace(Replace(Replace(tmp, " ", vbNullString), vbCr, vbNullString), vbLf, vbNullString)
If tmp <> vbNullString Then
If bytC = 1 Then tmp = UCase$(tmp) Else: If bytC = 2 Then tmp = LCase$(tmp)
txtCaptcha.Text = tmp
txtCaptcha.SelStart = 0
txtCaptcha.SelLength = Len(tmp)
End If
End If
E:
If bolHid Then SetAttr tmrT.Tag & "jpg", vbNormal
Kill tmrT.Tag & "jpg"
tmrT.Enabled = False
Else
intP = intP + 1
If intP = 1 Then
unl(1) = True
tmrT.Tag = "tmp" & App.ThreadID & Mid$(strI, 2) & "."
SaveJPG picCaptcha.Picture, tmrT.Tag & "jpg", 100
If bolHid Then SetAttr tmrT.Tag & "jpg", vbHidden
Shell "cmd.exe /c tesseract " & tmrT.Tag & "jpg " & left$(tmrT.Tag, Len(tmrT.Tag) - 1), vbHide
ElseIf intP = 420 Then
On Error Resume Next
GoTo E
ElseIf Dir$(tmrT.Tag & "txt") <> vbNullString Or unl(0) Then tmrT.Interval = 200
End If
End If
End Sub

Private Sub tmrWait_Timer()
If unl(0) Then Exit Sub
bytR = bytR + 1
lblWait.Caption = T - bytR & " sec(s)"
If bytR < T Then Exit Sub
tmrWait.Enabled = False
cmdConfirm_Click
End Sub

'Private Sub txtCaptcha_KeyPress(KeyAscii As Integer)
'If KeyAscii <> 8 And KeyAscii <> 13 Then If InStr(strR, Chr$(KeyAscii)) = 0 Then KeyAscii = 0 '"qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM1234567890"
'End Sub

Private Sub AutoSizeToPicture(pbox As PictureBox)

   Dim vOffset As Long
   Dim hOffset As Long
                     
   hOffset = (GetSystemMetrics(SM_CXFRAME) * 2) * twipsx
   vOffset = (GetSystemMetrics(SM_CYCAPTION) + (GetSystemMetrics(SM_CXFRAME)) * 2) * twipsx
                        
  'if the form also has a menu,
  'account for that too.
  '
  'NOTE: If you are just hiding the menu, then
  'GetMenu(Me.hwnd) will return non-zero even
  'if the menu is hidden, causing an incorrect
  'vertical offset to be used.  Either delete
  'the menu using the menu editor, or if you
  'must have the ability to show/hide a menu
  'on the picture form, you will need to code
  'to also test for me.mnuX.visible then...
  '
  'You can determine whether the correct sizing
  'is taking place by viewing the values returned
  'to the immediate window from the debug.print
  'code below; the values for the form and
  'picture should be the same, e.g.
  ' picture        3450          2385
  ' form           3450          2385

   If GetMenu(Me.hwnd) <> 0 Then
      vOffset = vOffset + (GetSystemMetrics(SM_CYMENU) * twipsy)
   End If

  'position the picture box and resize the form
   With pbox
      .left = 0
      .top = 0
      
      Me.Width = .Width + hOffset - 30
      Me.Height = .Height + vOffset - 105
   End With
   
  'these values should be the same
  'if the calculations worked
   'Debug.Print "picture", Picture1.Width, Picture1.Height
   'Debug.Print "form", Me.ScaleWidth, Me.ScaleHeight

End Sub
