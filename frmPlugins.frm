VERSION 5.00
Begin VB.Form frmPlugins 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Plugins"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4935
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAll 
      Caption         =   "&A"
      Height          =   195
      Left            =   5040
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   120
      Width           =   195
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&..."
      Height          =   255
      Left            =   4320
      TabIndex        =   10
      ToolTipText     =   "Change"
      Top             =   3000
      Width           =   495
   End
   Begin VB.TextBox txtInfo 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   4
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton cmdUnload 
      Caption         =   "&Unload"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   1680
      Width           =   735
   End
   Begin VB.ListBox lstNotLoaded 
      Height          =   1230
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   1815
   End
   Begin VB.ListBox lstPlugins 
      Height          =   1230
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox txtLocation 
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton cmdSettings 
      Caption         =   "&Settings"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CheckBox chkStartup 
      Caption         =   "&Load on startup"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtInfo 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   0
      Left            =   3150
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   1665
   End
   Begin VB.TextBox txtInfo 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   1
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox txtInfo 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   2
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox txtInfo 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   885
      Index           =   3
      Left            =   2160
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label8 
      Caption         =   "File:"
      Height          =   255
      Left            =   2160
      TabIndex        =   19
      Top             =   360
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "Location:"
      Height          =   255
      Left            =   2160
      TabIndex        =   18
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Version:"
      Height          =   255
      Left            =   2160
      TabIndex        =   17
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Author:"
      Height          =   255
      Left            =   2160
      TabIndex        =   16
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Command(s):"
      Height          =   255
      Left            =   2160
      TabIndex        =   15
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Description:"
      Height          =   255
      Left            =   2160
      TabIndex        =   14
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Not loaded:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "Double-click on some to load it."
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Loaded:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmPlugins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strLocation As String, strRL As String, strP As String, strPl As String, strNL As String, strC As String, strRg As String
Dim strP1 As String, bolL As Boolean, bolL1 As Boolean
Dim frmSettings As Object

Private Sub chkStartup_Click()
If bolL Then Exit Sub
'Dim p As Integer: p = InStr(InStr(strP, "|" & lstPlugins.ListIndex & "|"), strP, vbLf) - 1
If chkStartup.Value = 1 Then strP = Replace(strP, "|" & lstPlugins.ListIndex & "|", "|" & lstPlugins.ListIndex & "S|") Else: strP = Replace(strP, "|" & lstPlugins.ListIndex & "S|", "|" & lstPlugins.ListIndex & "|")
'p = InStr(p, strP, "S")
'strP = Left$(strP, p - 1) & Mid$(strP, p + 1)
'Else: 'strP = Left$(strP, p) & "S" & Mid$(strP, p + 1)
'End If
End Sub

Private Sub cmdAll_Click()
Dim i As Byte
If strRg <> vbNullString Then i = 1
If lstPlugins.ListCount < 2 + i Then Exit Sub
If MsgBox("Unload all plugins?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
Do While lstPlugins.ListCount > i
lstPlugins.ListIndex = i
cmdUnload_Click
Loop
End Sub

Private Sub cmdBrowse_Click()
Dim strT As String, strT1 As String
If strLocation <> vbNullString Then strT = strLocation Else: strT = frmMain.strLastPath
strT1 = strLocation
strLocation = GetFolder(Me.hWnd, strT, "Select plugins folder")
If strLocation = vbNullChar Then strLocation = strT1: Exit Sub
If Right$(strLocation, 1) <> "\" Then strLocation = strLocation & "\"
'If strLocation = strT1 Then Exit Sub
Dim strL As String, strL1 As String
If txtInfo(4).Text <> vbNullString Then
ExtrF lstPlugins.ListIndex, strL
strL1 = Split(strP, "|" & strL & "|")(0)
strL1 = Mid$(strL1, InStrRev(vbLf & strL1, vbLf, Len(strL1) - 1))
txtInfo(4).Text = frmMain.get_relative_path_to(strL1 & strL, , strLocation)
End If
strP1 = vbNullString
lstNotLoaded.Clear
ScanPlugins strLocation
strRL = frmMain.get_relative_path_to(strLocation, True)
FillNL
End Sub

Sub ExtrF(Index As Integer, strT As String, Optional bytP As Byte)
If bytP <> 1 Then
If InStr(strP, "|" & Index & "|") = 0 Then strT = Split(strP, "|" & Index & "S|")(0) Else: strT = Split(strP, "|" & Index & "|")(0)
Else: strT = Left$(strP1, InStr(strP1, "|" & Index & "|") - 1)
End If
strT = Mid$(strT, InStrRev(vbLf & strT, vbLf, Len(strT) - 1))
If bytP = 2 Then strT = Replace(strT, "|", vbNullString): Exit Sub
strT = Mid$(strT, InStr(strT, "|") + 1)
strT = Left$(strT, InStr(strT & "|", "|") - 1)
End Sub

Private Sub cmdSettings_Click()
frmSettings.Show vbModal
End Sub

Private Sub cmdUnload_Click()
Dim s() As String, i As Byte, pos As Integer, p As Integer ', strT1 As String
p = lstPlugins.ListIndex
Dim strT As String
ExtrF p, strT
On Error Resume Next
Set frmSettings = Nothing
Plugins.Remove strT & "/" & lstPlugins.list(p)
strC = Replace(strC, "|" & strT & "/" & lstPlugins.list(p) & "|" & Split(Split(strC, "|" & strT & "/" & lstPlugins.list(p) & "|")(1), vbLf)(0) & vbLf, vbNullString)
On Error GoTo 0
lstPlugins.RemoveItem p
pos = InStr(strP, "|" & p & "|")
If pos = 0 Then pos = InStr(strP, "|" & p & "S|")
'pos = InStr(vbLf & strP, vbLf & strT)
'strT1 = Split(Split(strP, strT & "|")(1), vbLf)(0)
'If strT1 <> vbNullString Then
's() = Split(strT1, "|")
'strP = Replace(strP, strT & "|" & strT1 & vbLf, vbNullString)
'Else: strP = Replace(strP, strT & vbLf, vbNullString)
'End If
'For i = 0 To UBound(s) - 1
'Plugins.Remove strT & "/" & lstPlugins.List(i)
'lstPlugins.RemoveItem p
'Next
'p = UBound(s)
strP = Replace(Replace(strP, "|" & p & "|", "|"), "|" & p & "S|", "|")
If InStr(strP, "|" & strT & "|" & vbLf) > 0 Then
Dim strT1 As String
strT1 = Split(strP, "|" & strT & "|")(0)
strT1 = Mid$(strT1, InStrRev(vbLf & strT1, vbLf, Len(strT1) - 1))
pos = InStr(vbLf & strP, vbLf & strT1 & "|" & strT & "|" & vbLf) + 1
strP = Replace(strP, strT1 & "|" & strT & "|" & vbLf, vbNullString)
If strLocation <> strT1 Then
strP1 = strP1 & strT1 & strT & "|" & lstNotLoaded.ListCount & "|" & vbLf
strT = strT1 & strT
UnloadLibrary strT
Else: UnloadLibrary strLocation & strT
End If
lstNotLoaded.AddItem frmMain.get_relative_path_to(strT, , strLocation)
SetListboxScrollbar1 lstNotLoaded
lstNotLoaded.ListIndex = lstNotLoaded.ListCount - 1
lstNotLoaded.SetFocus
Else
GoSub E
If p <= lstPlugins.ListCount - 1 Then lstPlugins.ListIndex = p Else: lstPlugins.ListIndex = lstPlugins.ListCount - 1
Exit Sub
End If
E:
If Len(strP) > pos Then
s() = Split(Mid$(strP, pos), "|")
For i = IIf(s(0) <> vbNullString, 2, 1) To UBound(s) - 1
If IsNumeric(Replace(s(i), "S", vbNullString)) Then strP = Replace(strP, "|" & s(i) & "|", "|" & Replace(s(i), "S", vbNullString) - 1 & IIf(InStr(s(i), "S") > 0, "S", vbNullString) & "|", , 1)
Next
End If
On Error Resume Next
Return
'Debug.Print strP & "{CRLF}" & vbCrLf & strPl & "{CRLF}" & vbCrLf & strNL & "{CRLF}" & vbCrLf & strC
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
If frmMain.chkOnTop.Checked Then SetTopMostWindow Me.hWnd, True
'del
'strLocation = strInitD & "Plugins\"
'ScanPlugins strLocation, True
'del
LoadLst lstPlugins
strPl = vbNullString
FillNL
End Sub

Private Sub FillNL()
Dim s() As String, i As Integer
s() = Split(strNL, vbLf)
For i = 0 To UBound(s) - 1
lstNotLoaded.AddItem s(i)
Next
strNL = vbNullString
SetListboxScrollbar1 lstNotLoaded
txtLocation.Text = strRL
txtLocation.SelStart = Len(txtLocation.Text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
For i = IIf(strRg <> vbNullString, 1, 0) To lstPlugins.ListCount - 1
strPl = strPl & lstPlugins.list(i) & vbLf
Next
For i = 0 To lstNotLoaded.ListCount - 1
strNL = strNL & lstNotLoaded.list(i) & vbLf
Next
Set frmSettings = Nothing
End Sub

Private Sub lstNotLoaded_Click()
If lstNotLoaded.ListIndex = -1 Or lstNotLoaded.ListCount = 0 Then Exit Sub
If strRg = vbNullString Then
C: If Not chkStartup.Enabled Then Exit Sub
ElseIf lstPlugins.ListIndex > 0 Then GoTo C
End If
bolL = True
chkStartup.Value = 0
bolL = False
chkStartup.Enabled = False
cmdSettings.Enabled = False
cmdUnload.Enabled = False
Dim i As Byte
For i = 0 To 4
txtInfo(i).Text = vbNullString
Next
lstPlugins.ListIndex = -1
End Sub

Private Sub lstNotLoaded_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And lstNotLoaded.ListIndex <> -1 Then lstNotLoaded_DblClick
If KeyCode <> 46 Then Exit Sub
If lstNotLoaded.ListIndex = -1 Or lstNotLoaded.ListCount = 0 Then Exit Sub
Dim strT As String
If InStr(strP1, "|" & lstNotLoaded.ListIndex & "|" & vbLf) > 0 Then
ExtrF lstNotLoaded.ListIndex, strT, 1
strP1 = Replace(strP1, strT & "|" & lstNotLoaded.ListIndex & "|" & vbLf, vbNullString)
End If
lstNotLoaded.RemoveItem lstNotLoaded.ListIndex
SetListboxScrollbar1 lstNotLoaded
End Sub

Private Sub lstPlugins_Click()
If lstPlugins.ListIndex = -1 Or lstPlugins.ListCount = 0 Or bolL Then Exit Sub
lstNotLoaded.ListIndex = -1
If lstPlugins.ListIndex = 0 And strRg <> vbNullString Then
cmdUnload.Enabled = False
chkStartup.Value = 1
chkStartup.Enabled = False
cmdSettings.Enabled = False
txtInfo(4).Text = strRg
txtInfo(0).Text = "rg1"
txtInfo(1).Text = ".NET 3.5"
txtInfo(2).Text = "MikiSoft"
txtInfo(3).Text = "This is a special plugin which relies on .NET Framework." & vbNewLine & "It's used like original RegEx function." & vbNewLine & "Example: rg1('[src]','regex','replacement',1,2) - while last three fields are optional (last two, from left to right: start - default 0, and count - default: result number)"
Exit Sub
End If
Dim Data(2) As String, i As Byte, strT As String
ExtrF lstPlugins.ListIndex, strT
Dim strT1 As String
strT1 = Split(strP, "|" & strT & "|")(0)
strT1 = Mid$(strT1, InStrRev(vbLf & strT1, vbLf, Len(strT1) - 1))
bolL = True
If InStr(strP, "|" & lstPlugins.ListIndex & "S|") > 0 Then chkStartup.Value = 1 Else: chkStartup.Value = 0
bolL = False
If strT1 <> strLocation Then txtInfo(4).Text = frmMain.get_relative_path_to(strT1 & strT, , strLocation) Else: txtInfo(4).Text = strT
txtInfo(4).SelStart = Len(txtInfo(4).Text)
On Error GoTo E
Dim tmpObj As IPluginInterface
Set tmpObj = Plugins.Item(strT & "/" & lstPlugins.list(lstPlugins.ListIndex))
Set frmSettings = Nothing
Set frmSettings = tmpObj.Info(Data)
Set tmpObj = Nothing
If frmSettings Is Nothing Then cmdSettings.Enabled = False Else: If TypeOf frmSettings Is Form Then cmdSettings.Enabled = True Else: cmdSettings.Enabled = False
For i = 0 To 2
txtInfo(i + 1).Text = Trim$(Data(i))
Next
strT = Split(Split(strC, "|" & strT & "/" & lstPlugins.list(lstPlugins.ListIndex) & "|")(1), vbLf)(0)
strT = Mid$(strT, 2, Len(strT) - 2)
txtInfo(0).Text = Replace(Replace(strT, " ", vbNullString), ",", ", ")
If chkStartup.Enabled Then Exit Sub
chkStartup.Enabled = True
cmdUnload.Enabled = True
Exit Sub
E:
MsgBox "There was some error!", vbExclamation
If bolL1 Then bolL1 = False Else: cmdUnload_Click
End Sub

Private Sub lstNotLoaded_DblClick()
If lstNotLoaded.ListIndex = -1 Or lstNotLoaded.ListCount = 0 Then Exit Sub
Dim i As Integer, strT(1) As String, strT1 As String
i = lstNotLoaded.ListIndex
If InStr(strP1, "|" & i & "|" & vbLf) > 0 Then
ExtrF i, strT(0), 1
strT(1) = Left$(strT(0), InStrRev(strT(0), "\"))
strT1 = strT(0)
strT(0) = Mid$(strT(0), Len(strT(1)) + 1)
Else
strT(0) = lstNotLoaded.list(lstNotLoaded.ListIndex)
strT(1) = strLocation
End If
'Dim l As Byte: l = frmMain.LoadRG(strT(1), strT(0))
'If l = 2 Then
'MsgBox "There was some error while loading plugin!", vbExclamation
'Exit Sub
'ElseIf l = 1 Then
'MsgBox "ok"
'Exit Sub
'End If
If Not LoadPlugin(strT(0), strT(1)) Then
frmMain.bolChk = True
lstNotLoaded.RemoveItem i
SetListboxScrollbar1 lstNotLoaded
Dim s() As String, pos As Integer
pos = InStr(strP1, "|" & i & "|")
If pos > 0 Then
s() = Split(Mid$(strP1, pos), "|")
Dim a As Byte
For a = 3 To UBound(s) - 1 Step 2
strP1 = Replace(strP1, "|" & s(a) & "|", "|" & s(a) - 1 & "|", , 1)
Next
If strT1 <> vbNullString Then strP1 = Replace(strP1, strT1 & "|" & i & "|" & vbLf, vbNullString)
End If
Else: MsgBox "There was some error while loading this plugin!", vbExclamation
End If
End Sub

Private Function LoadPlugin(strFN As String, Optional strL As String) As Boolean

  ' Try load plugins
  Dim clsid() As GUID
  Dim names() As String
  Dim cnt   As Long
  Dim l As Boolean
  
  On Error GoTo ERROR_LOADING
  
  cnt = 0
  
  If strL = vbNullString Then strL = strLocation
  
  ' Get all co-classes in dll (this support several plugins in one dll)
  If GetAllCoclasses(strL & strFN, clsid(), names(), cnt) Then
  
  strP = strP & strL & "|" & strFN
      ' New error handler
      On Error Resume Next
      
      Dim tmpObj As IPluginInterface
      
      Do While cnt
      
          Set tmpObj = CreateObjectEx(strL & strFN, clsid(cnt - 1))
          
          ' Object created and support IPluginInterface
          If Not tmpObj Is Nothing Then
              l = True
              ' Add it to list
              Plugins.add tmpObj, strFN & "/" & names(cnt - 1)
              strC = strC & "|" & strFN & "/" & names(cnt - 1) & "|," & TrimComma(tmpObj.Startup(frmMain)) & "," & vbLf
              lstPlugins.AddItem names(cnt - 1)
              strP = strP & "|" & lstPlugins.ListCount - 1
              bolL = True
              lstPlugins.ListIndex = lstPlugins.ListCount - 1
              bolL = False
              bolL1 = True
              lstPlugins_Click
              If Not bolL1 Then
              txtInfo(4).Text = vbNullString
              lstPlugins.RemoveItem lstPlugins.ListCount - 1
              Plugins.Remove Plugins.count
              Else: bolL1 = False
              End If
              lstPlugins.SetFocus
          'Else: GoTo ERROR_LOADING
          End If
          cnt = cnt - 1
          
      Loop
      
      On Error GoTo -1
      UnloadLibrary strL & strFN
  End If
  
E:
  Set tmpObj = Nothing
  If l Then If InStr(Split(strP, "|" & strFN)(1), "|") = 0 Then strP = Replace(strP, strL & "|" & strFN, vbNullString) Else: strP = strP & "|" & vbLf Else: strP = Replace(strP, strL & "|" & strFN, vbNullString): LoadPlugin = True
  Exit Function

ERROR_LOADING:
  LoadPlugin = True
  GoTo E
End Function
