VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   5145
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Text            =   "rg(rpl(captcha(a)))+rpl(rpl('a(b)','b','a'),'a''b')+'''a'"
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Text2.Text = ProcessString(Text1.Text)
End Sub

Function ProcessString(ByVal strInp As String) As String
Dim strT(1) As String, lngC(1) As Long
strT(1) = "result"
lngC(0) = Len(strInp)
If lngC(0) = 0 Then Exit Function
Do
Select Case Mid$(strInp, lngC(0), 1)
Case ")"
lngC(1) = Len(strInp) - lngC(0) + 1
Do
strT(0) = FindCommRev(strInp, lngC(0))
If strT(0) = vbNullString Then lngC(0) = lngC(0) - 1: Exit Do
strInp = Replace(strInp, strT(0), "'" & Replace(strT(1), "'", "''") & "'")
Debug.Print strInp
lngC(0) = lngC(0) - Len(strT(0)) + Len(strT(1)) + 2
Loop Until Left$(Right$(strInp, lngC(1)), 1) <> ")"
Case "'"
lngC(0) = lngC(0) - 1
lngC(1) = FindCRev(strInp, lngC(0))
ProcessString = Replace(Mid$(strInp, lngC(1) + 1, lngC(0) - lngC(1)), "''", "'") & ProcessString
lngC(0) = lngC(1) - 2
Case "+": lngC(0) = lngC(0) - 1
Case Else
N:
lngC(1) = InStrRev(strInp, "+", lngC(0)) + 1
If InStr("-0123456789", Mid$(strInp, lngC(1), 1)) > 0 Then
strT(0) = Mid$(strInp, lngC(1), lngC(0) - lngC(1) + 1)
If IsNumeric(ProcessString) Then ProcessString = ProcessString + Val(strT(0)) Else: ProcessString = strT(0) & ProcessString
lngC(0) = lngC(1) - 2
If lngC(0) = 1 Then GoTo N
Else: Exit Function
End If
End Select
Loop Until lngC(0) < 2
End Function

Function FindCommRev(strExp As String, Optional lngS As Long) As String
Dim lngC(4) As Long
If lngS = 0 Then lngS = Len(strExp)
lngC(1) = lngS
lngC(2) = lngS
Do
If lngC(3) > 0 Then lngC(4) = lngC(3) + lngC(1) - 1
lngC(0) = lngC(2)
lngC(1) = InStrRev(strExp, "'", lngC(2)) - 1
If lngC(1) < 1 Then Exit Do
lngC(2) = FindCRev(strExp, lngC(1)) - 1
lngC(1) = lngC(1) + 2
lngC(3) = InStr(Mid$(strExp, lngC(1), lngC(0) - lngC(1) + 1), ")")
Loop Until InStrRev(Mid$(strExp, lngC(1), lngC(0) - lngC(1) + 1), "(") > 0 Or lngC(2) < 1
If lngC(2) > 0 Then
If lngC(4) = 0 Then
lngC(4) = InStr(Left$(strExp, lngC(0)), ")")
If lngC(4) = 0 Then
lngC(4) = lngS
strExp = Left$(strExp, lngS) & ")" & Mid$(strExp, lngS + 1)
lngC(4) = lngC(4) + 1
End If
End If
ElseIf lngC(4) = 0 Then lngC(4) = lngC(3) + lngC(1) - 1
End If
If lngC(2) = -1 Then lngC(0) = InStrRev(strExp, "(", lngC(1)) Else: lngC(0) = InStrRev(strExp, "(", lngC(0))
If lngC(0) = 0 Then Exit Function
Do
lngC(0) = lngC(0) - 1
If lngC(0) = 1 Then Exit Do
Loop Until InStr("+(,", Mid$(strExp, lngC(0) - 1, 1)) > 0
FindCommRev = Mid$(strExp, lngC(0), lngC(4) - lngC(0) + 1)
End Function

Private Function FindCRev(strExp As String, lngC As Long) As Long
Dim bytC As Byte
FindCRev = InStrRev(strExp, "'", lngC)
Do While FindCRev > 1
bytC = 1
Do While Mid$(strExp, FindCRev - 1, 1) = "'"
FindCRev = FindCRev - 1
bytC = bytC + 1
If FindCRev = 1 Then Exit Do
Loop
If bytC Mod 2 <> 0 Then Exit Function
FindCRev = InStrRev(strExp, "'", FindCRev - 1)
Loop
End Function

Private Sub Form_Load()
Command1_Click
End Sub
