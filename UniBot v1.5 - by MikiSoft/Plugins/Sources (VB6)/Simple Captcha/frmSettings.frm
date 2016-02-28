VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Simple Captcha settings"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2895
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkCurr 
      Caption         =   "Onl&y for this session"
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
      Left            =   120
      TabIndex        =   18
      ToolTipText     =   "Uncheck if you want to save settings for next time, ie. to write settings file."
      Top             =   3000
      Width           =   1815
   End
   Begin VB.TextBox txtCustom 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   5
      Top             =   650
      Width           =   1935
   End
   Begin VB.OptionButton optBoth 
      Caption         =   "&Both"
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
      Left            =   1850
      TabIndex        =   9
      Top             =   1260
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.OptionButton optLower 
      Caption         =   "L&ower"
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
      Left            =   980
      TabIndex        =   8
      Top             =   1260
      Width           =   855
   End
   Begin VB.OptionButton optUpper 
      Caption         =   "&Upper"
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
      Left            =   120
      TabIndex        =   7
      Top             =   1260
      Width           =   855
   End
   Begin VB.CheckBox chkSym 
      Caption         =   "&Symbols"
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
      Left            =   1920
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin VB.CheckBox chkLett 
      Caption         =   "&Letters"
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
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.CheckBox chkDigits 
      Caption         =   "&Digits"
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
      Left            =   1070
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox txtMax 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   13
      ToolTipText     =   "Leave zero for unlimited."
      Top             =   1620
      Width           =   495
   End
   Begin VB.TextBox txtMin 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   11
      ToolTipText     =   "Leave zero for unlimited."
      Top             =   1620
      Width           =   495
   End
   Begin VB.TextBox txtTimeout 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   15
      ToolTipText     =   "In second(s). Leave zero for unlimited."
      Top             =   2000
      Width           =   495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
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
      Height          =   1215
      Left            =   2280
      TabIndex        =   19
      Top             =   2040
      Width           =   495
   End
   Begin VB.CheckBox chkRem 
      Caption         =   "&Remember on each call"
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
      Left            =   120
      TabIndex        =   17
      ToolTipText     =   "Rememeber changes after each calling of this plugin."
      Top             =   2720
      Width           =   2025
   End
   Begin VB.CheckBox chkOCR 
      Caption         =   "Use &Tesseract OCR"
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
      Left            =   120
      TabIndex        =   16
      Top             =   2370
      Width           =   1695
   End
   Begin VB.Label lblC 
      Caption         =   "&Custom:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   650
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Allowed case:"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1020
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Allowed characters:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "M&ax. l:"
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
      Left            =   1680
      TabIndex        =   12
      Top             =   1620
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "M&in. length:"
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
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Minimum characters length of captcha response."
      Top             =   1620
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "R&esponse timeout:"
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
      Left            =   120
      TabIndex        =   14
      Top             =   2000
      Width           =   1455
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bolNoLett As Boolean
Public bolNoDigits As Boolean
Public bolNoSym As Boolean
Public strCustom As String
Public bytCase As Byte
Public bytMin As Byte
Public bytMax As Byte
Public bytTimeout As Byte
Public bolOCR As VbTriState
Public bolNoRem As Boolean
Public bolNoCurr As Boolean

Private Sub chkSym_Click()
If txtCustom.Text = vbNullString And chkSym.Value = 0 And chkDigits.Value = 0 And chkLett.Value = 0 Then chkSym.Value = 1: Exit Sub
ChkCust
End Sub

Private Sub chkDigits_Click()
If txtCustom.Text = vbNullString And chkDigits.Value = 0 And chkLett.Value = 0 And chkSym.Value = 0 Then chkDigits.Value = 1: Exit Sub
ChkCust
End Sub

Private Sub chkLett_Click()
If txtCustom.Text = vbNullString And chkSym.Value = 0 And chkDigits.Value = 0 And chkLett.Value = 0 Then chkLett.Value = 1: Exit Sub
ChkCust
End Sub

Private Sub cmdOK_Click()
bolNoLett = Not CBool(chkLett.Value)
bolNoDigits = Not CBool(chkDigits.Value)
bolNoSym = Not CBool(chkSym.Value)
strCustom = txtCustom.Text
If optUpper.Value Then bytCase = 1 Else: If optLower.Value Then bytCase = 2 Else: bytCase = 0
If txtMin.Text < 255 Then bytMin = txtMin.Text Else: bytMin = 255
If txtMax.Text < 255 Then bytMax = txtMax.Text Else: bytMax = 255
If bytMax > 0 And bytMax < bytMin Then bytMax = bytMin
If txtTimeout.Text < 255 Then bytTimeout = txtTimeout.Text Else: bytTimeout = 255
If bolOCR <> vbUseDefault Then bolOCR = CBool(chkOCR.Value)
bolNoRem = Not CBool(chkRem.Value)
bolNoCurr = Not CBool(chkCurr.Value)
Unload Me
End Sub

Private Sub Form_Load()
chkLett.Value = CInt(Not bolNoLett) * (-1)
chkDigits.Value = CInt(Not bolNoDigits) * (-1)
chkSym.Value = CInt(Not bolNoSym) * (-1)
txtCustom.Text = strCustom
If bytCase = 1 Then optUpper.Value = True Else: If bytCase = 2 Then optLower.Value = True
txtMin.Text = bytMin
txtMax.Text = bytMax
txtTimeout.Text = bytTimeout
If bolOCR <> vbUseDefault Then chkOCR.Value = CInt(bolOCR) * (-1) Else: chkOCR.Enabled = False
chkRem.Value = CInt(Not bolNoRem) * (-1)
chkCurr.Value = CInt(Not bolNoCurr) * (-1)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub optLower_Click()
ChkCust
End Sub

Private Sub optUpper_Click()
ChkCust
End Sub

Private Sub txtCustom_Change()
ChkCust
txtCustom.SelStart = Len(txtCustom.Text)
End Sub

Private Sub ChkCust()
If txtCustom.Text = vbNullString Then Exit Sub
Dim bytT As Byte: If optUpper.Value Then bytT = 1 Else: If optLower.Value Then bytT = 2
txtCustom.Text = ProcStr(txtCustom.Text, Not CBool(chkLett.Value), Not CBool(chkDigits.Value), Not CBool(chkSym.Value), bytT)
End Sub

Private Sub txtCustom_KeyPress(KeyAscii As Integer)
If KeyAscii <> 1 Then Exit Sub
txtCustom.SelStart = 0
txtCustom.SelLength = Len(txtCustom.Text)
End Sub

Private Sub txtMin_Change()
CheckText txtMin
End Sub

Private Sub txtMax_Change()
CheckText txtMax
End Sub

Private Sub txtTimeout_Change()
CheckText txtTimeout
End Sub

Private Function CheckText(obj As TextBox)
On Error Resume Next
If obj.Text = vbNullString Then
obj.Text = 0
obj.SelStart = 1
obj.Tag = vbNullString
Exit Function
End If
If IsNumeric(obj.Text) Then
If obj.Text < 0 Then
Dim intS As Integer
If obj.SelStart > 0 Then intS = obj.SelStart - 1 Else: intS = obj.SelStart
obj.Text = obj.Text * (-1)
obj.SelStart = intS
End If
obj.Tag = CLng(obj.Text)
End If
intS = obj.SelStart
obj.Text = obj.Tag
obj.SelStart = intS
End Function
