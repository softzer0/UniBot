VERSION 5.00
Begin VB.Form frmS 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "captcha9kw settings"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2460
      Width           =   3135
   End
   Begin VB.CheckBox chkNoSave 
      Caption         =   "&Don't save"
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
      Left            =   2160
      TabIndex        =   17
      ToolTipText     =   "Use current settings only for this session"
      Top             =   2125
      Width           =   1095
   End
   Begin VB.TextBox txtMaxL 
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
      Top             =   1680
      Width           =   495
   End
   Begin VB.CheckBox chkRem 
      Caption         =   "Remember each &chng."
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
      ToolTipText     =   "Rememeber changes after each calling of this plugin."
      Top             =   2125
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Default options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   3135
      Begin VB.CheckBox chkNum 
         Caption         =   "&Num."
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
         TabIndex        =   3
         ToolTipText     =   "Numeric"
         Top             =   240
         Width           =   705
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
         Left            =   1380
         TabIndex        =   13
         Top             =   865
         Width           =   495
      End
      Begin VB.TextBox txtPrior 
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
         Left            =   2685
         TabIndex        =   11
         ToolTipText     =   "Leave zero for unlimited."
         Top             =   1220
         Width           =   275
      End
      Begin VB.CheckBox chkNoSpace 
         Caption         =   "No spac&e"
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
         Left            =   2040
         TabIndex        =   9
         Top             =   865
         Width           =   975
      End
      Begin VB.CheckBox chkOCR 
         Caption         =   "&OCR"
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
         Left            =   2400
         TabIndex        =   6
         Top             =   240
         Width           =   705
      End
      Begin VB.CheckBox chkSym 
         Caption         =   "with s&ymbols"
         Enabled         =   0   'False
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
         Left            =   1560
         TabIndex        =   8
         Top             =   550
         Width           =   1300
      End
      Begin VB.CheckBox chkMath 
         Caption         =   "Mat&h"
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
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox chkCase 
         Caption         =   "Case &sensitive"
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
         Top             =   550
         Width           =   1335
      End
      Begin VB.CheckBox chkPhrase 
         Caption         =   "&Phrase"
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
         Left            =   840
         TabIndex        =   4
         ToolTipText     =   "Captcha contains two or more words."
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Ma&ximum length:"
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
         ToolTipText     =   "Maximum length"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "M&inimum length:"
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
         TabIndex        =   12
         Top             =   865
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "P&riority:"
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
         Left            =   2040
         TabIndex        =   10
         Top             =   1220
         Width           =   615
      End
   End
   Begin VB.TextBox txtKey 
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
      MaxLength       =   32
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "&API key:"
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
      Width           =   615
   End
End
Attribute VB_Name = "frmS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strKey As String
Public bytMin As Byte
Public bytMaxL As Byte
Public bytCase As Boolean
Public bytPrior As Byte
Public bolNum As Boolean
Public bolMath As Boolean
Public bolPhrase As Boolean
Public bolOCR As Boolean
Public bolNoSpace As Boolean
Public bolNoRem As Boolean
Public bolSave As Boolean

Function RplKey(strT As String) As String
If Len(strT) >= 5 And Len(strT) <= 50 Then If Not strT Like "*[!A-Za-z0-9]*" Then RplKey = strT
End Function

Private Sub chkCase_Click()
chkSym.Enabled = CBool(chkCase.Value)
End Sub

Private Sub cmdOK_Click()
strKey = RplKey(Replace(Replace(Replace(txtKey.Text, vbCr, vbNullString), vbLf, vbNullString), vbTab, vbNullString))
If strKey = vbNullString Then If MsgBox("You have entered key of invalid length so it won't be saved! Continue?", vbExclamation + vbYesNo) = vbNo Then Exit Sub
If txtPrior.Text < 20 Then bytPrior = txtPrior.Text Else: bytPrior = 20
ChngM txtMin, txtMaxL, bytMin, bytMaxL, 3
If chkCase.Value = 1 Then bytCase = chkCase.Value + chkSym.Value
bolMath = CBool(chkMath.Value)
bolNum = CBool(chkNum.Value)
bolPhrase = CBool(chkPhrase.Value)
bolOCR = CBool(chkOCR.Value)
bolNoSpace = CBool(chkNoSpace.Value)
bolNoRem = Not CBool(chkRem.Value)
bolSave = Not CBool(chkNoSave.Value)
Unload Me
End Sub

Private Sub ChngM(txt1 As TextBox, txt2 As TextBox, bytM As Byte, bytM1 As Byte, bytDMin As Byte)
If txt2.Text < 255 Then bytM1 = txt2.Text Else: bytM1 = 255
If txt1.Text < 255 Then
If txt1.Text > 0 Then bytM = txt1.Text Else: If bytM1 > 0 Then bytM = bytM1 Else: bytM = bytDMin
Else: bytM = 255
End If
If bytM1 > 0 And bytM1 < bytM Then bytM1 = bytM
End Sub

Private Sub Form_Load()
txtKey.Text = strKey
txtMin.Text = bytMin
txtMaxL.Text = bytMaxL
txtPrior.Text = bytPrior
If bytCase >= 1 Then
chkCase.Value = 1
If bytCase = 2 Then chkSym.Value = 1
End If
chkMath.Value = CInt(bolMath) * (-1)
chkPhrase.Value = CInt(bolPhrase) * (-1)
chkNum.Value = CInt(bolNum) * (-1)
chkOCR.Value = CInt(bolOCR) * (-1)
chkNoSpace.Value = CInt(bolNoSpace) * (-1)
chkRem.Value = CInt(Not bolNoRem) * (-1)
chkNoSave.Value = CInt(Not bolSave) * (-1)
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

Private Sub txtMin_Change()
CheckText txtMin
End Sub

Private Sub txtMaxL_Change()
CheckText txtMaxL
End Sub

Private Sub txtPrior_Change()
CheckText txtPrior
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub txtKey_KeyPress(KeyAscii As Integer)
If KeyAscii = 1 Then
KeyAscii = 0
txtKey.SelStart = 0
txtKey.SelLength = Len(txtKey.Text)
ElseIf KeyAscii = 8 Or KeyAscii = 22 Or KeyAscii = 3 Or KeyAscii = 24 Then Exit Sub
ElseIf Chr(KeyAscii) Like "[A-Za-z0-9]" Then Exit Sub
Else: KeyAscii = 0
End If
End Sub
