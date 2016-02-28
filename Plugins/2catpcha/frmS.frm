VERSION 5.00
Begin VB.Form frmS 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "2captcha settings"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   3855
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
      TabIndex        =   24
      Top             =   3430
      Width           =   3615
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
      Left            =   2640
      TabIndex        =   23
      ToolTipText     =   "Use current settings only for this session"
      Top             =   3110
      Width           =   1095
   End
   Begin VB.CheckBox chkRem 
      Caption         =   "Remember &chng. on each call"
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
      TabIndex        =   22
      ToolTipText     =   "Rememeber changes after each calling of this plugin."
      Top             =   3110
      Width           =   2415
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
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   3615
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
         Left            =   2160
         TabIndex        =   19
         ToolTipText     =   "Captcha contains two or more words."
         Top             =   1880
         Width           =   855
      End
      Begin VB.TextBox txtAfter 
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
         Left            =   3120
         TabIndex        =   21
         ToolTipText     =   "In seconds. Maximum 255. Leave zero to disable storing."
         Top             =   2160
         Width           =   375
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   495
         Left            =   1200
         TabIndex        =   9
         Top             =   840
         Width           =   2295
         Begin VB.OptionButton optCyr 
            Caption         =   "C&yrillic"
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
            Left            =   960
            TabIndex        =   11
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton optLat 
            Caption         =   "La&tin"
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
            TabIndex        =   10
            Top             =   0
            Width           =   735
         End
         Begin VB.OptionButton optNot 
            Caption         =   "N&ot specified"
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
            Index           =   1
            Left            =   0
            TabIndex        =   12
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.CheckBox chkMath 
         Caption         =   "Has mat&h"
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
         TabIndex        =   18
         Top             =   1640
         Width           =   1095
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
         Left            =   1200
         TabIndex        =   6
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton optLett 
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
         Left            =   2160
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton optNum 
         Caption         =   "&Numeric"
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
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optNot 
         Caption         =   "N&ot specified"
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
         Index           =   0
         Left            =   2160
         TabIndex        =   7
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
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
         Left            =   2160
         TabIndex        =   17
         Top             =   1400
         Width           =   1335
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
         Left            =   1380
         TabIndex        =   16
         ToolTipText     =   "Leave zero for unlimited."
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtWait 
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
         TabIndex        =   14
         ToolTipText     =   "In second(s). Maximum 255."
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "S&tore for reporting, then purge all after:"
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
         TabIndex        =   20
         ToolTipText     =   "Clear all unused stored captcha IDs on each:"
         Top             =   2175
         Width           =   3015
      End
      Begin VB.Label Label5 
         Caption         =   "Language:"
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
         TabIndex        =   8
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Cap. type:"
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
         ToolTipText     =   "Captcha type"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "&Max. resp. time:"
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
         TabIndex        =   15
         ToolTipText     =   "Maximum response time"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "&Response wait:"
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
         TabIndex        =   13
         Top             =   1440
         Width           =   1215
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
      Width           =   2895
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
Public bytType As Byte
Public bytLang As Byte
Public bytWait As Byte
Public bytMax As Byte
Public bolCase As Boolean
Public bolMath As Boolean
Public bolPhrase As Boolean
Public bolNoRem As Boolean
Public bytAfter As Byte
Public bolSave As Boolean

Function RplKey(strT As String) As String
If Len(strT) <> 32 Then Exit Function
Dim i As Byte
For i = 1 To 32
If InStr("qwertyuiopasdfghjklzxcvbnm0123456789", Mid$(strT, i, 1)) = 0 Then Exit Function
Next
RplKey = strT
End Function

Private Sub cmdOK_Click()
strKey = RplKey(Replace(Replace(Replace(txtKey.Text, vbCr, vbNullString), vbLf, vbNullString), vbTab, vbNullString))
If strKey = vbNullString Then If MsgBox("You have entered key of invalid length so it won't be saved! Continue?", vbExclamation + vbYesNo) = vbNo Then Exit Sub
If optNot(0).Value Then
bytType = 0
ElseIf optNum.Value Then bytType = 1
ElseIf optLett.Value Then bytType = 2
ElseIf optBoth.Value Then bytType = 3
End If
If optNot(1).Value Then
bytLang = 0
ElseIf optLat.Value Then bytLang = 1
ElseIf optCyr.Value Then bytLang = 2
End If
If txtMax.Text < 255 Then bytMax = txtMax.Text Else: bytMax = 255
If txtWait.Text < 255 Then
If txtWait.Text > 0 Then bytWait = txtWait.Text Else: If bytMax > 0 Then bytWait = bytMax Else: bytWait = 5
Else: bytWait = 255
End If
If bytMax > 0 And bytMax < bytWait Then bytMax = bytWait
bolCase = CBool(chkCase.Value)
bolMath = CBool(chkMath.Value)
bolPhrase = CBool(chkPhrase.Value)
If txtAfter.Text < 255 Then bytAfter = txtAfter.Text Else: bytAfter = 255
bolNoRem = Not CBool(chkRem.Value)
bolSave = Not CBool(chkNoSave.Value)
Unload Me
End Sub

Private Sub Form_Load()
txtKey.Text = strKey
Select Case bytType
Case 0: optNot(0).Value = True
Case 1: optNum.Value = True
Case 2: optLett.Value = True
Case 3: optBoth.Value = True
End Select
Select Case bytLang
Case 0: optNot(1).Value = True
Case 1: optLat.Value = True
Case 2: optCyr.Value = True
End Select
txtWait.Text = bytWait
txtMax.Text = bytMax
txtAfter.Text = bytAfter
chkCase.Value = CInt(bolCase) * (-1)
chkMath.Value = CInt(bolMath) * (-1)
chkPhrase.Value = CInt(bolPhrase) * (-1)
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

Private Sub txtAfter_Change()
CheckText txtAfter
End Sub

Private Sub txtMax_Change()
CheckText txtMax
End Sub

Private Sub txtWait_Change()
CheckText txtWait
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
ElseIf InStr("qwertyuiopasdfghjklzxcvbnm0123456789", Chr(KeyAscii)) > 0 Then Exit Sub
Else: KeyAscii = 0
End If
End Sub
