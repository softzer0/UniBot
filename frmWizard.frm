VERSION 5.00
Begin VB.Form frmWizard 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Wizard"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5175
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Frame fra2 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Index           =   1
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   4935
         Begin VB.TextBox txtFName 
            BackColor       =   &H8000000F&
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
            Left            =   3840
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   180
            Width           =   1095
         End
         Begin VB.CommandButton cmdDisplay 
            Caption         =   "&Display"
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
            Index           =   2
            Left            =   1440
            TabIndex        =   32
            Top             =   2880
            Width           =   2055
         End
         Begin VB.VScrollBar VScroll1 
            Enabled         =   0   'False
            Height          =   2295
            Left            =   4680
            Max             =   4
            TabIndex        =   27
            Top             =   480
            Width           =   255
         End
         Begin VB.PictureBox picO 
            Height          =   2295
            Left            =   0
            ScaleHeight     =   2235
            ScaleWidth      =   4635
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   480
            Width           =   4695
            Begin VB.PictureBox picI 
               BorderStyle     =   0  'None
               Height          =   2295
               Index           =   0
               Left            =   0
               ScaleHeight     =   2295
               ScaleWidth      =   4695
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   0
               Width           =   4695
               Begin VB.TextBox txtValues 
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   615
                  Index           =   0
                  Left            =   1080
                  MultiLine       =   -1  'True
                  ScrollBars      =   3  'Both
                  TabIndex        =   31
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   150
               End
               Begin VB.ListBox lstValues 
                  Height          =   645
                  Index           =   0
                  ItemData        =   "frmWizard.frx":0000
                  Left            =   840
                  List            =   "frmWizard.frx":0007
                  MultiSelect     =   2  'Extended
                  TabIndex        =   30
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   135
               End
               Begin VB.ComboBox cmbValues 
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   0
                  Left            =   120
                  Style           =   2  'Dropdown List
                  TabIndex        =   29
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   615
               End
            End
         End
         Begin VB.Label Label7 
            Caption         =   "Form:"
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
            Left            =   3360
            TabIndex        =   24
            Top             =   180
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Enter the data that you would fill:"
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
            TabIndex        =   23
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame fra2 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Index           =   0
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Visible         =   0   'False
         Width           =   4935
         Begin VB.ListBox lstFN 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2010
            Index           =   1
            Left            =   0
            TabIndex        =   15
            Top             =   600
            Width           =   2175
         End
         Begin VB.CommandButton cmdDisplay 
            Caption         =   "&Display"
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
            Left            =   1440
            TabIndex        =   21
            Top             =   2880
            Width           =   2055
         End
         Begin VB.ListBox lstFN 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2010
            Index           =   0
            Left            =   2760
            TabIndex        =   18
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label4 
            Caption         =   "OR"
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
            Left            =   2340
            TabIndex        =   16
            Top             =   1605
            Width           =   375
         End
         Begin VB.Label lblU 
            Caption         =   "Pick a &URL for login:"
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
            TabIndex        =   14
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lblU 
            Caption         =   "Pick a &form for login:"
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
            Left            =   2760
            TabIndex        =   17
            Top             =   360
            Width           =   1575
         End
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   34
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "&Finish"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   35
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Back"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   33
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CommandButton cmdDisplay 
         Caption         =   "Display"
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
         Left            =   1440
         TabIndex        =   19
         Top             =   405
         Width           =   3375
      End
      Begin VB.TextBox txtLoc 
         BackColor       =   &H8000000F&
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   120
         Width           =   3375
      End
      Begin VB.OptionButton optExtract 
         Caption         =   "Extract some data"
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
         Left            =   1440
         TabIndex        =   12
         Top             =   2520
         Width           =   2055
      End
      Begin VB.OptionButton optFN 
         Caption         =   "Navigate to some page"
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
         Index           =   1
         Left            =   1440
         TabIndex        =   11
         Top             =   2160
         Width           =   2055
      End
      Begin VB.OptionButton optFN 
         Caption         =   "Submit some form"
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
         Index           =   0
         Left            =   1440
         TabIndex        =   10
         Top             =   1800
         Width           =   2055
      End
      Begin VB.OptionButton optLogin 
         Caption         =   "Login"
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
         Left            =   1440
         TabIndex        =   9
         Top             =   1440
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "Current location:"
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
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "What do you want to do next?"
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
         Left            =   0
         TabIndex        =   8
         Top             =   960
         Width           =   4935
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4935
      Begin VB.TextBox txtSite 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Welcome to wizard for creating your own configuration!"
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
         Left            =   0
         TabIndex        =   4
         Top             =   120
         Width           =   4935
      End
      Begin VB.Label Label3 
         Caption         =   "For what site are you creating configuration? Enter here its address (e.g. http://site.com):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   3
         Top             =   960
         Width           =   3555
      End
   End
   Begin VB.Label lblS 
      Alignment       =   2  'Center
      Caption         =   "Step 1: Web site"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strUA As String
Dim bytS(1) As Byte, strSrc As String, rh As WinHttp.WinHttpRequest, ctlValues() As VB.Control, bytCP As Byte, lngTotP As Long
Dim strName() As String, strURLData() As String, varHeaders() As Variant, varStrings() As Variant, colForms As Collection

Private Sub cmbValues_Click(Index As Integer)
cmbValues(Index).ToolTipText = cmbValues(Index).Text
End Sub

Private Sub cmdBack_Click()
BackNext -1
End Sub

Private Sub cmdNext_Click()
BackNext 1
End Sub

Private Sub BackNext(intS As Integer)
'del
fra2(1).Visible = True
fra(2).Visible = True
GoTo t
'del
bytS(0) = bytS(0) + intS
lblS.Tag = lblS.Caption
lblS.Caption = vbNullString
'On Error GoTo E1 'enable
Dim s() As String, i As Integer, a As Integer, strT As String, bolE As Boolean
Select Case bytS(0)
Case 0
If intS = -1 Then
cmdBack.Enabled = False
lblS.Caption = "Web site"
GoSub ClrD
End If
Case 1
If intS = 1 Then
frmMain.lblStatus.Caption = "Opening web site..."
frmMain.lblStatus.Refresh
SetSite
strSrc = GetPage(txtSite.Text)
If Len(Mid$(strSrc, InStr(strSrc & vbNewLine & vbNewLine, vbNewLine & vbNewLine) + 4)) > 0 Then
txtLoc.Text = txtSite.Text
optLogin.FontBold = False
frmMain.lblStatus.Caption = "Analyzing page..."
frmMain.lblStatus.Refresh
ExtrUF "a", "href", 1
ExtrUF "form", "action"
optLogin.Enabled = optFN(0).Tag <> vbNullString Or optFN(1).Tag <> vbNullString
optFN(0).Enabled = optFN(0).Tag <> vbNullString
optFN(1).Enabled = optFN(1).Tag <> vbNullString
Select Case True
Case optLogin.Value And Not optLogin.Enabled, optFN(0).Value And Not optFN(0).Enabled, optFN(1).Value And Not optFN(1).Enabled
If Not optLogin.Enabled Then
If Not optFN(0).Enabled Then
If Not optFN(1).Enabled Then optExtract.Value = True Else: optFN(1).Value = True
Else: optFN(0).Value = True
End If
Else: optLogin.Value = True
End If
End Select
cmdBack.Enabled = True
Else
Er:
frmMain.lblStatus.Caption = "Error!"
MsgBox "There is some error with accessing the site!" & vbNewLine & vbNewLine & "Either it is:" & vbNewLine & "- wrong URL" & vbNewLine & "- blank page" & vbNewLine & "- connection problem", vbExclamation
bolE = True
End If
Else: GoSub ClrL
End If
frmMain.lblStatus.Caption = "Idle..."
If bolE Then GoTo E
lblS.Caption = "The next thing"
If intS = -1 Then cmdFinish.Enabled = False
Case 2
If intS = -1 Then
cmdFinish.Enabled = False
bolE = True
End If
Select Case True
Case optLogin.Value And Not optLogin.FontBold, optFN(0).Value, optFN(1).Value
fra2(1).Visible = False
fra2(0).Visible = True
End Select
If optLogin.Value Then
If Not optLogin.FontBold Then
lblS.Caption = "Locate login"
If intS = 1 Then
For a = 0 To 1
GoSub FillLst
Next
lblU(1).Caption = "Pick a &URL for login:"
lblU(0).Caption = "Pick a &form for login:"
If lstFN(1).ListCount > 0 And lstFN(0).ListCount > 0 Then
If Not lstFN(0).Visible Or Not lstFN(1).Visible Then
lblU(0).Left = 2760
lstFN(0).Left = 2760
For a = 0 To 1
lstFN(a).Width = 2175
lblU(a).Visible = True
lstFN(a).Visible = True
Next
End If
Else
If lstFN(1).ListCount > 0 Then a = 1 Else: a = 0
GoSub PosLst
End If
End If
Else
fra2(0).Visible = False
If Left$(optLogin.Tag, 1) = vbCr Then
strT = Mid$(optLogin.Tag, 2)
GoTo C
Else
fra2(1).Visible = False
'...
End If
End If
ElseIf optFN(0).Value Then
lblS.Caption = "Select form"
If intS = 1 Then
lblU(0).Caption = "Pick a &form:"
a = 0
GoSub FillLst
End If
ElseIf optFN(1).Value Then
lblS.Caption = "Select URL"
If intS = 1 Then
lblU(1).Caption = "Pick a &URL:"
a = 1
GoSub FillLst
End If
End If
Case 3
bolE = True
If optFN(0).Value Or optLogin.Value And lstFN(0).ListIndex > -1 Then
strT = lstFN(0).list(lstFN(0).ListIndex)
C:
lngTotP = 0
frmMain.lblStatus.Caption = "Examining form..."
frmMain.lblStatus.Refresh
txtFName.Text = strT
strT = "<form " & Split(strSrc, "<form ", , vbTextCompare)(colForms.Item(strT))
strT = Left$(strT, InStr(1, strT & "</form>", "</form>", vbTextCompare) + 6)
t: strT = LoadFile(App.path & "\test.html") 'del
picI(bytCP).Tag = 0
cmdDisplay(2).Tag = strT
lblS.Caption = "Fill in the data"
fra2(1).Visible = True
Dim ctl As VB.TextBox, lbl As VB.Label, p As Long, R As String, b As Integer, C As Integer, strN As String, strT1 As String, bolT As Boolean, lngP(3) As Long, ch As String, strT2 As String
lngP(0) = 1
lngP(1) = InStr(lngP(0), strT, "<input ", vbTextCompare)
lngP(2) = InStr(lngP(0), strT, "<select ", vbTextCompare)
lngP(3) = InStr(lngP(0), strT, "<textarea ", vbTextCompare)
Do While lngP(1) > 0 Or lngP(2) > 0 Or lngP(3) > 0
ReDim Preserve ctlValues(a)
If lngP(1) > 0 And (lngP(2) > lngP(1) Or lngP(2) = 0) And (lngP(3) > lngP(1) Or lngP(3) = 0) Then
strT1 = Mid$(strT, lngP(1))
GoSub ExtrNC
If strN <> vbNullString Then
Select Case ExtrParam(strT1, "type")
'Case "hidden"
'...
'GoTo N2
Case "text", "password"
Set ctlValues(a) = Controls.add("VB.TextBox", "txtValue" & a, picI(bytCP))
ctlValues(a).MaxLength = Val(ExtrParam(strT1, "maxlength"))
Case "checkbox"
Set ctlValues(a) = Controls.add("VB.CheckBox", "chkValue" & a, picI(bytCP))
ctlValues(a).Caption = "Checked"
Case "radio"
Select Case AddCmb(ExtrParam(strT1, "value"), strN, a, b, R)
Case vbTrue: bolT = True
Case vbUseDefault: GoTo N1
End Select
Case Else: GoTo N1
End Select
GoSub CrN
a = a + 1
End If
N1: lngP(0) = lngP(1) + 7
ElseIf lngP(2) > 0 And (lngP(3) > lngP(2) Or lngP(3) = 0) Then
strT1 = Mid$(strT, lngP(2))
GoSub ExtrNC
If strN <> vbNullString Then
s() = Split(Left$(strT1, FStr(strT1, "</select>", ch, vbTextCompare)), "<option ", , vbTextCompare)
If UBound(s) > 0 Then
For i = 1 To UBound(s)
strT2 = Mid$(s(i), InStr(s(i), "=") + 1, 1)
If InStr("""'", strT2) = 0 Then strT2 = vbNullString
s(i) = Mid$(s(i), FStr(s(i), ">", strT2) + 1)
AddCmb StripTags(Left$(s(i), InStr(s(i) & "<", "<") - 1)), strN, a, b, R
If i = 1 Then bolT = True: GoSub CrN
Next
a = a + 1
End If
End If
lngP(0) = lngP(2) + 8
Else
'...
lngP(0) = lngP(3) + 10
End If
If lngP(1) > 0 Then lngP(1) = InStr(lngP(0), strT, "<input ", vbTextCompare)
If lngP(2) > 0 Then lngP(2) = InStr(lngP(0), strT, "<select ", vbTextCompare)
If lngP(3) > 0 Then lngP(3) = InStr(lngP(0), strT, "<textarea ", vbTextCompare)
Debug.Print strN
'If strN = "sound[]" Then
'Debug.Print
'End If
Loop
'...
'frmMain.lblStatus.Caption = "Idle..." 'enable
'Debug.Print VScroll1.LargeChange, VScroll1.SmallChange
End If
Case 4
If intS = 1 Then cmdFinish.Enabled = True
GoSub ClrD
'Case n - 1
'If intS = -1 Then
'cmdNext.Enabled = True
'cmdNext.Default = True
'End If
'...
'Case n
'cmdNext.Enabled = False
'cmdFinish.Default = True
End Select
N:
Exit Sub 'del
lblS.Caption = "Step " & Split(Split(lblS.Tag, " ")(1), ":")(0) + intS & ": " & lblS.Caption
If Not bolE Then
fra(bytS(0) - intS).Visible = False
fra(bytS(0)).Visible = True
End If
Exit Sub
E1:
Me.Enabled = True
frmMain.lblStatus.Caption = "Idle..."
MsgBox "There is some unexpected error!", vbCritical
E:
lblS.Caption = lblS.Tag
bytS(0) = bytS(0) - 1
Exit Sub
ClrD:
Set colForms = Nothing
optFN(0).Tag = vbNullString
optFN(1).Tag = vbNullString
If bytS(0) = 0 Then Return
ClrL:
For a = 0 To 1
lstFN(a).ToolTipText = vbNullString
lstFN(a).Clear
Next
Return
FillLst:
s() = Split(optFN(a).Tag, vbLf)
For i = 0 To UBound(s) - 1
lstFN(a).AddItem s(i)
SetListboxScrollbar1 lstFN(a)
Next
If optLogin.Value Then lstFN(0).ListIndex = 0: Return
PosLst:
lstFN(a).ListIndex = 0
If lstFN(a).Visible And lstFN(a).Width = 4935 Then Return
If a = 1 Then
lstFN(0).Visible = False
lblU(0).Visible = False
Else
lstFN(1).Visible = False
lblU(1).Visible = False
End If
lstFN(a).Left = 0
lstFN(a).Width = 4935
lblU(a).Visible = True
lblU(a).Left = 0
lstFN(a).Visible = True
Return
CrN:
p = CLng(C) * 385 + 100
C = C + 1
If bolT Then
ctlValues(a).Move 2450, p, 2125
If bytCP > 0 Then Set ctlValues(a).Container = picI(bytCP)
bolT = False
Else: ctlValues(a).Move 2450, p, 2125, 285
End If
If ctlValues(a).Top + ctlValues(a).Height + 100 > 245745 Then
bytCP = bytCP + 1
Load picI(bytCP)
C = 1
p = 100
Set ctlValues(a).Container = picI(bytCP)
ctlValues(a).Top = p
picI(bytCP).Width = picI(0).Width
picI(bytCP - 1).Height = picI(bytCP - 1).Height - 100
picI(bytCP).Top = picI(bytCP - 1).Top + picI(bytCP - 1).Height
picI(bytCP).Tag = picI(bytCP).Top
picI(bytCP).Visible = True
End If
picI(bytCP).Height = ctlValues(a).Top + ctlValues(a).Height + 100
lngTotP = picI(bytCP).Top + picI(bytCP).Height
ctlValues(a).Tag = strN
ctlValues(a).Font.Name = "Tahoma"
ctlValues(a).TabIndex = 31
ctlValues(a).Visible = True
Set ctl = Controls.add("VB.TextBox", "txtName" & a, picI(bytCP))
ctl.Move 100, p, 2125, 285
ctl.Font.Name = "Tahoma"
ctl.Locked = True
ctl.BackColor = picI(bytCP).BackColor
ctl.Text = strN
Dim p1 As Byte
strT2 = ExtrParam(strT1, "placeholder")
p1 = FStr(strT1, ">", ch) + 1
Dim strN1(1) As String
On Error Resume Next
strN1(0) = StripTags(Mid$(strT1, p1, FStr(Left$(strT1, InStr(p1, strT1, ">")), "</label>", ch, vbTextCompare) - p1))
On Error GoTo 0 'repl: E1
strN1(1) = ExtrParam(Left$(strT1, p1), "id")
If strN1(1) <> vbNullString Then If strN1(1) <> Replace(ctl.Text, "[]", vbNullString, , 1) Then ctl.Text = strN1(1) & ", " & ctl.Text
If strN1(0) = vbNullString Then
strN1(0) = vbNullChar
Nm:
strN1(1) = "<label for=" & ch & strN1(1) & ch & ">"
strN1(1) = Split(strT & strN1(1), strN1(1), , vbTextCompare)(1)
If strN1(1) <> vbNullString Then strN1(1) = StripTags(Left$(strN1(1), FStr(strN1(1) & "</label>", "</label>", ch, vbTextCompare) - 1))
If strN1(0) = vbNullChar And strN1(1) = vbNullString Then strN1(0) = vbNullString: strN1(1) = strN: GoTo Nm
If strN1(1) <> vbNullString Then If strT2 <> vbNullString Then strT2 = strT2 & "; " & strN1(1) Else: strT2 = strN1(1)
Else: If strT2 <> vbNullString Then strT2 = strT2 & "; " & strN1(0) Else: strT2 = strN1(0)
End If
strN1(0) = vbNullString
strN1(1) = vbNullString
If strT2 <> vbNullString Then ctl.Text = strT2 & " (" & ctl.Text & ")": strT2 = vbNullString
ctl.TabIndex = 31
ctl.Visible = True
Set lbl = Controls.add("VB.Label", "lbl" & a, picI(bytCP))
lbl.Move 2275, p + 20, 135, 255
lbl.Caption = "="
lbl.Visible = True
With VScroll1
If Not .Enabled Then If picI(bytCP).Height > picO.Height Then .Enabled = True ': .Tag = 0
If .Enabled Then
Dim vl(1) As Integer
'If a <= 32767 Then
.Max = a
If lngTotP < 2850 Then '10 * 285
.SmallChange = Round(CLng(.Max) * 3 / 10)
.LargeChange = .Max
ElseIf lngTotP > 7125 Then '25 * 285
If .SmallChange > 1 Then
vl(0) = Round(CLng(.SmallChange) * 9.5 / 10)
If vl(0) > 0 Then .SmallChange = vl(0)
End If
If .LargeChange > 1 Then
vl(1) = Round(CLng(.LargeChange) * 9.5 / 10)
If vl(1) > 0 Then .LargeChange = vl(1)
End If
Else
.SmallChange = Round(CLng(.Max) * 2 / 10)
.LargeChange = Round(CLng(.Max) * 8 / 10)
End If
'ElseIf .Tag > -1 Then
'.Tag = .Tag + 1
'If .Tag = Round(CLng(.SmallChange) * 9.5 / 10) Then
'vl(0) = Round(CLng(.SmallChange) * 9.5 / 10)
'vl(1) = Round(CLng(.LargeChange) * 9.5 / 10)
'If vl(0) > 0 Or vl(1) > 0 Then
'.Tag = 0
'If vl(0) > 0 Then .SmallChange = vl(0)
'If vl(1) > 0 Then .LargeChange = vl(1)
'Else: .Tag = -1
'End If
'End If
'End If
End If
End With
Return
ExtrNC:
strN = ExtrParam(strT1, "name")
For i = 0 To UBound(ctlValues) - 1
If ctlValues(i).Tag = strN Then strN = vbNullString: Exit For
Next
If strN <> vbNullString Then
ch = Mid$(strT1, InStr(strT1, "=") + 1, 1)
If InStr("""'", ch) = 0 Then ch = vbNullString
End If
Return
End Sub

Private Function StripTags(ByVal strContent As String) As String
On Error Resume Next
Dim mString As String
Dim mStartPos As Integer, mEndPos As Integer
mStartPos = InStr(strContent, "<")
mEndPos = InStr(strContent, ">")
StripTags = strContent
Do While mStartPos <> 0 And mEndPos <> 0 And mEndPos > mStartPos
      mString = Mid$(StripTags, mStartPos, mEndPos - mStartPos + 1)
      StripTags = Replace(StripTags, mString, vbNullString)
      mStartPos = InStr(StripTags, "<")
      mEndPos = InStr(StripTags, ">")
Loop
StripTags = Replace(StripTags, "&nbsp;", " ")
StripTags = Replace(StripTags, "&amp;", "&")
StripTags = Replace(StripTags, "&quot;", "'")
StripTags = Replace(StripTags, "&#", "#")
StripTags = Replace(StripTags, "&lt;", "<")
StripTags = Replace(StripTags, "&gt;", ">")
StripTags = Replace(StripTags, "%20", " ")
StripTags = Replace(StripTags, vbTab, vbNullString)
StripTags = Trim$(StripTags)
Do While Left$(StripTags, 1) = vbCr Or Left$(StripTags, 1) = vbLf
      StripTags = Mid$(StripTags, 2)
Loop
Do While Right$(StripTags, 1) = vbCr Or Right$(StripTags, 1) = vbLf
      StripTags = Left$(StripTags, Len(StripTags) - 1)
Loop
End Function

Private Function AddCmb(strT As String, strN As String, a As Integer, b As Integer, R As String) As VbTriState
If InStr(R, """'" & strN & """'") = 0 Then
If cmbValues(0).Tag <> vbNullString Then
b = cmbValues.count
Load cmbValues(b)
Else: b = 0
End If
R = R & """'" & strN & """'" & a & ","
cmbValues(b).AddItem strT
cmbValues(b).ListIndex = 0
Set ctlValues(a) = cmbValues(b)
AddCmb = vbTrue
Else
ctlValues(Split(Split(R, """'" & strN & """'")(1), ",")(0)).AddItem strT
Set ctlValues(a) = cmbValues(b)
AddCmb = vbUseDefault
End If
End Function

Private Sub ExtrUF(strTag As String, strAttr As String, Optional a As Integer)
Dim s() As String, strT As String, strT1 As String, i As Integer
s() = Split(strSrc, "<" & strTag & " ", , vbTextCompare)
If a = 0 Then Set colForms = New Collection
For i = 1 To UBound(s)
strT = ExtrParam(s(i), strAttr, a)
If strT = vbNullString Then GoTo N
Select Case True
Case LCase$(Left$(strT, 8)) = "https://", LCase$(Left$(strT, 7)) = "http://", Left$(strT, 1) = "/"
strT1 = Mid$(strT, InStr(strT, "//") + 2)
C:
If InStr(Replace(Replace(Replace(strT1, "www.", vbNullString, , 1, vbTextCompare), "http://", vbNullString, , 1, vbTextCompare), "https://", vbNullString, , 1, vbTextCompare), Replace(Replace(Replace(txtSite.Text, "www.", vbNullString, , 1, vbTextCompare), "http://", vbNullString, , 1, vbTextCompare), "https://", vbNullString, , 1, vbTextCompare)) = 1 Then
If Left$(strT, 1) <> "/" Then If Left$(strT, 8) = "https://" Then If Left$(txtSite.Text, InStr(txtSite.Text, "/")) <> "https://" Then txtSite.Text = "https:" & Mid$(txtSite.Text, InStr(txtSite.Text, "/"))
If Left$(strT1, InStr(strT1 & "/", "/")) <> Mid$(txtSite.Text, InStr(txtSite.Text, "//") + 2) Then txtSite.Text = Left$(txtSite.Text, InStr(txtSite.Text, "//") + 1) & Left$(strT1, InStr(strT1 & "/", "/"))
strT = Mid$(strT1, InStr(strT1, "/") + 1)
Else: GoTo N
End If
Case Left$(Replace(strT, "/", vbNullString), 1) = "#": GoTo N
Case InStr(strT, ":") > InStr(strT, "/"): GoTo N
Case Left$(strT, 1) = "/": strT1 = Mid$(strT, 2): GoTo C
End Select
If a = 0 Or strT <> vbNullString Then
If Right$(strT, 1) = "/" Then strT = Left$(strT, Len(strT) - 1)
If Not optLogin.FontBold And a = 1 Or Left$(optLogin.Tag, 1) <> vbCr And a = 0 Then ChkL strT, a 'enable
If a = 0 Then
On Local Error GoTo -1
On Local Error GoTo N
colForms.add i, strT
End If
If InStr(vbLf & optFN(a).Tag, vbLf & strT & vbLf) = 0 Then optFN(a).Tag = optFN(a).Tag & strT & vbLf
N:
If a = 0 Then On Local Error GoTo 0
End If
Next
End Sub

Private Function ChkL(strT As String, Optional a As Integer)
Dim strT1 As String: strT1 = Replace(Replace(Replace(strT, "_", vbNullString), "-", vbNullString), ".", vbNullString)
optLogin.FontBold = InStr(1, strT1, "login", vbTextCompare) > 0 Or InStr(1, strT1, "signin", vbTextCompare) > 0
If optLogin.FontBold Then If a = 0 Then optLogin.Tag = vbCr & strT Else: optLogin.Tag = strT
End Function

Private Function ExtrParam(strS As String, ByVal strParam As String, Optional a As Integer) As String
strParam = strParam & "="
Dim strT As String, p(1) As Long, p2 As Byte
p(0) = InStr(strS, "'")
p(1) = InStr(strS, """")
p2 = InStr(strS, ">")
If p2 < p(1) And p2 < p(0) And p2 > 0 Then
ExtrParam = Left$(strS, p2)
Else
If p(0) > p(1) Or p(0) = 0 Then strT = """" Else: If p(0) > 0 Then strT = "'" Else: Exit Function
ExtrParam = Left$(strS, FindSep(strS, , ">", strT))
End If
p(0) = FStr(ExtrParam, strParam, strT)
If p(0) = 0 Then ExtrParam = vbNullString: Exit Function
ExtrParam = Mid$(ExtrParam, p(0) + Len(strParam))
If InStr("""'", Left$(ExtrParam, 1)) = 0 Then
p(0) = InStr(ExtrParam, vbTab)
p(1) = InStr(ExtrParam, " ")
p2 = InStr(ExtrParam, ">")
If p(0) = 0 And p(1) = 0 And p2 = 0 Then ExtrParam = vbNullString: Exit Function
If p(1) > 0 And (p(0) > p(1) Or p(0) = 0) And (p2 > p(1) Or p2 = 0) Then ExtrParam = Left$(ExtrParam, p(1) - 1) Else: If p(0) > 0 And (p2 > p(0) Or p2 = 0) Then ExtrParam = Left$(ExtrParam, p(0) - 1) Else: ExtrParam = Left$(ExtrParam, p2 - 1)
Else: ExtrParam = Left$(Mid$(ExtrParam, 2), InStr(2, ExtrParam, Left$(ExtrParam, 1)) - 2)
End If
If a <> 1 Then Exit Function
If Left$(ExtrParam, 1) = "." Then ExtrParam = Mid$(ExtrParam, 2)
If Left$(ExtrParam, 1) = "/" Then ExtrParam = Mid$(ExtrParam, 2)
End Function

Private Function FStr(strS As String, strT As String, strC As String, Optional comp As Integer = vbBinaryCompare) As Long
If strC <> vbNullString Then FStr = FindSep(strS, , strT, strC, comp) Else: FStr = InStr(1, strS, strT, comp)
End Function

Private Function GetPage(strSite As String) As String
On Error GoTo E
Dim strT(1) As String
R:
strT(1) = strSite & vbLf
rh.Open "GET", strSite
rh.SetRequestHeader "User-Agent", strUA
rh.Send
On Error Resume Next
strT(0) = rh.GetResponseHeader("Location")
On Error GoTo 0
On Error GoTo E
If strT(0) <> vbNullString Then
If InStr(vbLf & strT(1) & vbLf, vbLf & strT(0) & vbLf) > 0 Then Exit Function
strSite = strT(0)
strT(0) = vbNullString
If bytS(0) = 1 Then txtSite.Text = strSite: SetSite
GoTo R
End If
GetPage = rh.Status & " " & rh.StatusText & vbNewLine & rh.GetAllResponseHeaders
On Error Resume Next
GetPage = GetPage & frmMain.FromCPString(rh.ResponseBody, CP_UTF8)
Dim s() As String, i As Integer
s() = Split(GetPage, "<!--")
GetPage = s(0)
For i = 1 To UBound(s)
GetPage = GetPage & Mid$(s(i), InStr(s(i) & "-->", "-->") + 3)
Next
E:
End Function

Private Sub SetSite()
If LCase$(Left$(txtSite.Text, 8)) <> "https://" And LCase$(Left$(txtSite.Text, 7)) <> "http://" Then
txtSite.Text = "http://" & Left$(txtSite.Text, InStr(txtSite.Text & "/", "/"))
Else: txtSite.Text = Left$(txtSite.Text & "/", InStr(InStr(txtSite.Text, "//") + 2, txtSite.Text & "/", "/"))
End If
End Sub

Private Sub cmdDisplay_Click(Index As Integer)
Dim strT As String
Select Case Index
Case 0
strT = Mid$(strSrc, InStr(strSrc, vbNewLine & vbNewLine) + 4)
Dim s() As String, i As Integer
s() = Split("href="",src="",background="",url("",url('", ",")
For i = 0 To UBound(s)
strT = Replace(strT, s(i) & "//", StrReverse(s(i)) & "http://")
strT = Replace(strT, s(i) & "/", s(i) & txtSite.Text)
strT = Replace(strT, s(i) & "./", s(i) & txtSite.Text)
strT = Replace(strT, s(i) & "../", s(i) & txtSite.Text & "../")
strT = Replace(strT, s(i) & "http", s(i) & "1=""http")
strT = Replace(strT, s(i), s(i) & txtSite.Text)
strT = Replace(strT, StrReverse(s(i)), s(i))
Next
frmB.strSrc = strT
GoTo C1
Case 1
If lstFN(0).ListIndex > -1 Then
strT = lstFN(0).list(lstFN(0).ListIndex)
C:
strT = "<form " & Split(strSrc, "<form ", , vbTextCompare)(colForms.Item(strT))
frmB.strSrc = Left$(strT, InStr(1, strT & "</form>", "</form>", vbTextCompare) + 6)
End If
C1:
On Error Resume Next
Load frmB
If Not frmB.bolU Then Exit Sub
frmB.bolU = False
Unload frmB
Case 2
frmB.strSrc = cmdDisplay(2).Tag
GoTo C1
End Select
End Sub

Private Sub Form_Load()
'SetTopMostWindow Me.hWnd, True 'If frmMain.chkOnTop.Checked Then SetTopMostWindow Me.hWnd, True 'enable
Set rh = New WinHttp.WinHttpRequest
rh.Option(4) = 13056
rh.Option(WinHttpRequestOption_EnableRedirects) = False
rh.Option(12) = True
bytS(0) = 0
bytS(1) = 0
ReDim strURLData(0)
ReDim varStrings(0)
ReDim varHeaders(0)
ReDim strName(0)
varStrings(0) = Array("ua" & vbLf & strUA)
'del
'txtSite.Text = "http://feelingsurf.fr" '"https://www.ljuska.org"
'cmdNext_Click
'cmdNext_Click
'lstFN(0).ListIndex = 0
Me.Show
cmdNext_Click
'del
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set rh = Nothing
On Error Resume Next
Kill Environ("tmp") & "\UniBot.html"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me: Exit Sub
If ActiveControl Is Nothing Then Exit Sub
Debug.Print ActiveControl.Name, ActiveControl.TabIndex 'del
Const ASC_CTRL_A As Integer = 1

    ' See if this is Ctrl-A.
    If KeyAscii = ASC_CTRL_A Then
    KeyAscii = 0
        ' The user is pressing Ctrl-A. See if the
        ' active control is a TextBox.
        If TypeOf ActiveControl Is TextBox Then 'Or TypeOf ActiveControl Is UniTextBox
            ' Select the text in this control.
            ActiveControl.SelStart = 0
            ActiveControl.SelLength = Len(ActiveControl.Text)
        End If
    End If
End Sub

Private Sub lstFN_Click(Index As Integer)
If lstFN(Index).ListIndex = -1 Then Exit Sub
If Index = 0 Then lstFN(1).ListIndex = -1 Else: lstFN(0).ListIndex = -1
lstFN(Index).ToolTipText = lstFN(Index).list(lstFN(Index).ListIndex)
End Sub

Private Sub txtSite_Change()
If bytS(0) = 0 Then If Len(txtSite.Text) = 0 Then cmdNext.Enabled = False Else: If Not cmdNext.Enabled Then cmdNext.Enabled = True
End Sub

Private Sub VScroll1_Change()
ChngSC
End Sub

Private Sub VScroll1_Scroll()
ChngSC
End Sub

Private Sub ChngSC()
Dim lngT As Long, i As Byte
lngT = (VScroll1.Value / VScroll1.Max) * (lngTotP - picO.Height)
For i = 0 To bytCP
picI(i).Top = picI(i).Tag - lngT
picI(i).Refresh
Next
End Sub
