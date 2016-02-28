VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   705
   ClientWidth     =   4215
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
   MaxButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear log"
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   2100
      TabIndex        =   2
      Top             =   120
      Width           =   1985
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save log"
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1985
   End
   Begin VB.ListBox lstLog 
      Height          =   3180
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save output"
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   6615
      Width           =   1985
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear output"
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   2100
      TabIndex        =   5
      Top             =   6615
      Width           =   1985
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Default         =   -1  'True
      Height          =   255
      Left            =   2700
      TabIndex        =   6
      Top             =   7210
      Width           =   1395
   End
   Begin prjUB.UniTextBox txtOutput 
      Height          =   3015
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   5318
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      Text            =   ""
      MultiLine       =   -1  'True
      Locked          =   -1  'True
      Scrollbar       =   3
   End
   Begin VB.Timer tmrQ 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3600
      Top             =   6600
   End
   Begin VB.Timer tmrU 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1
      Left            =   2640
      Top             =   6600
   End
   Begin VB.Timer tmrI 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1
      Left            =   2880
      Top             =   6600
   End
   Begin VB.Timer tmrW 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1000
      Left            =   3360
      Top             =   6600
   End
   Begin VB.Label lblS1 
      Caption         =   "Status:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   6960
      Width           =   495
   End
   Begin VB.Label lblStatus 
      Caption         =   "Idle..."
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   6960
      Width           =   3375
   End
   Begin VB.Label lbl1 
      Caption         =   "UniBot"
      DragIcon        =   "frmM.frx":0000
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1815
      TabIndex        =   8
      Top             =   7210
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Bot created with"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   7210
      Width           =   1695
   End
   Begin VB.Menu adv 
      Caption         =   "Window"
      Begin VB.Menu chkOnTop 
         Caption         =   "Always on top"
         Shortcut        =   {F1}
      End
      Begin VB.Menu cmdMintoTray 
         Caption         =   "Minimize to tray"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'UniBot Stand-alone Application

Option Explicit

Private Plugins As New Collection

Private Const strDigit = "0123456789"
Private Const strLett = "qwertyuiopasdfghjklzxcvbnm"
Private Const strULett = "QWERTYUIOPASDFGHJKLZXCVBNM"
Private Const strSym = "~`!@#$%^&*()-=_+[]\{}|;':"",./<>?"

Dim Hash As New MD5Hash

Private Const GW_OWNER       As Long = 4

Private Const IMAGE_ICON     As Long = 1

Private Const ICON_SMALL     As Long = 0
Private Const ICON_BIG       As Long = 1

Private Const LR_DEFAULTSIZE As Long = &H40
Private Const LR_SHARED      As Long = &H8000

Private Const SM_CXICON      As Long = 11
Private Const SM_CYICON      As Long = 12
Private Const SM_CXSMICON    As Long = 49
Private Const SM_CYSMICON    As Long = 50

Private Const WM_SETICON     As Long = &H80

Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal uCmd As Long) As Long
Private Declare Function LoadImageA Lib "user32.dll" (ByVal hInst As Long, ByVal lpszName As Long, Optional ByVal uType As Long, Optional ByVal cxDesired As Long, Optional ByVal cyDesired As Long, Optional ByVal fuLoad As Long) As Long
Private Declare Function SendMessageA Lib "user32.dll" (ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private hWndOwner As Long

Private Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal _
        dest As Long, ByVal src As Long, ByVal Length As Long) As Long

Private Declare Function PathRelativePathTo Lib "shlwapi.dll" Alias "PathRelativePathToA" (ByVal pszPath As String, ByVal pszFrom As String, ByVal dwAttrFrom As Long, ByVal pszTo As String, ByVal dwAttrTo As Long) As Long
Private Const MAX_PATH As Long = 260
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80

Private Const LB_SETHORIZONTALEXTENT = &H194

Private Declare Function FindMimeFromData Lib "Urlmon.dll" ( _
    ByVal pBC As Long, _
    ByVal pwzUrl As Long, _
    ByVal pBuffer As Long, _
    ByVal cbSize As Long, _
    ByVal pwzMimeProposed As Long, _
    ByVal dwMimeFlags As Long, _
    ByRef ppwzMimeOut As Long, _
    ByVal dwReserved As Long _
) As Long
Private Const FMFD_DEFAULT As Long = &H0
Private Const FMFD_URLASFILENAME  As Long = &H1
Private Const FMFD_ENABLEMIMESNIFFING  As Long = &H2
Private Const FMFD_IGNOREMIMETEXTPLAIN  As Long = &H4
Private Const FMFD_SERVERMIME  As Long = &H8
Private Const FMFD_RESPECTTEXTPLAIN  As Long = &H10
Private Const FMFD_RETURNUPDATEDIMGMIMES  As Long = &H20
Private Const S_OK          As Long = 0&
Private Const E_FAIL        As Long = &H80000008
Private Const E_INVALIDARG  As Long = &H80000003
Private Const E_OUTOFMEMORY As Long = &H80000002

Private Declare Function lstrlen Lib "Kernel32.dll" Alias "lstrlenW" ( _
    ByVal lpString As Long _
) As Long

Private Declare Sub CoTaskMemFree Lib "ole32.dll" ( _
    ByVal pv As Long _
)

Private Declare Function WideCharToMultiByte Lib "Kernel32.dll" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long _
) As Long

Private Declare Function MultiByteToWideChar Lib "Kernel32.dll" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cbMultiByte As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long _
) As Long

Private Declare Function SetCurrentDirectoryA Lib "kernel32" (ByVal lpPathName As String) As Long

Private Const CP_ACP        As Long = 0          ' Default ANSI code page.
Private Const CP_UTF8       As Long = 65001      ' UTF8.
Private Const CP_UTF16_LE   As Long = 1200       ' UTF16 - little endian.
Private Const CP_UTF16_BE   As Long = 1201       ' UTF16 - big endian.
Private Const CP_UTF32_LE   As Long = 12000      ' UTF32 - little endian.
Private Const CP_UTF32_BE   As Long = 12001      ' UTF32 - big endian.

Private Declare Function CreateActCtx Lib "kernel32" Alias "CreateActCtxA" (ByRef pActCtx As ACTCTX_) As Long
Private Declare Sub ReleaseActCtx Lib "kernel32" (ByVal hActCtx As Long)
Private Declare Function ActivateActCtx Lib "kernel32" (ByVal hActCtx As Long, ByRef lpCookie As Long) As Boolean
Private Declare Function DeactivateActCtx Lib "kernel32" (ByVal dwFlags As Long, ByVal ulCookie As Long) As Boolean
Private Const INVALID_HANDLE_VALUE = -1
Private Const ACTCTX_FLAG_RESOURCE_NAME_VALID = 8&
Private Type ACTCTX_
    cbSize As Long
    dwFlags As Long
    lpSource As String
    wProcessorArchitecture As Integer
    wLangId As Integer
    lpAssemblyDirectory As String
    lpResourceName As String
    lpApplicationName As String
    hModule As Long
End Type

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) _
        As Long
Private WithEvents SystemTray As clsInTray
Attribute SystemTray.VB_VarHelpID = -1

Private Const Comms As String = ",rg,rpl,num,dech,dec,enc,u,l,b64,md5,"
Private WithEvents rh As cAsyncRequests
Attribute rh.VB_VarHelpID = -1
Dim strCmd As String, strPath(1) As String, bolAb As Boolean, bolEx As Boolean, strPlC As String, bytSh() As Byte, strPO As String, strInitD As String, strLastPath As String, strDrP As String
Dim bolNoRetry As Boolean, bytTimeout As Byte, bytThreads As Byte, bytSubThr As Byte, bytDelay As Byte, bytMaxR As Byte
Dim intAfter As Integer, bolHours As Boolean, strTemplate0 As String, strTemplate1 As String, bytTOrigin0 As Byte, bytTOrigin1 As Byte, bolNoEach As Boolean, intLogMax As Integer, intOutMax As Integer, bolColl As Boolean, bytPlgUse As Byte, bolUnl As Boolean, bolLO(1) As Boolean
Dim strC As String, bytActive As Byte, intSubT As Integer, intLTmr(1) As Integer, intTmrCount As Integer, bytOrigin As Byte, bolTmp As Boolean, bytSilent As Byte, bolMT As Boolean, bolSkipErr As Boolean, bolRg As Boolean, datCompl As Date
Dim bytIC As Byte, bytLimit As Byte, strURLData() As String, strHeaders() As String, strStrings() As String, strIf() As String, strWait() As String, intGoto() As Integer
Dim colSrc As Collection, colStr As Collection, colPubStr As Collection, colMax As Collection, colInput As Collection, colMaxR As Collection, colCurrO As Collection
Public bolDebug As Boolean

Private Sub Form_QueryUnlaod(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 And Not bolDebug Then If MsgBox("Are you sure?", vbExclamation + vbYesNo) = vbNo Then Cancel = 1 'If cmdStart.Caption = "Stop" Then
End Sub

Private Sub lbl1_DragDrop(Source As Control, x As Single, y As Single)
If Source Is lbl1 Then Shell "cmd.exe /c START http://unibot.boards.net", vbHide
End Sub

Private Sub lbl1_DragOver(Source As Control, x As Single, y As Single, State As Integer)
If State = vbLeave Then lbl1.Drag vbEndDrag
End Sub

Private Sub lbl1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
lbl1.Drag vbBeginDrag
End Sub

Sub addLog(txt As String, Optional D As Boolean)
txt = "[" & Now & "] " & txt
If D Then txt = "DEBUG: " & txt
If lstLog.ListCount = intLogMax Then lstLog.RemoveItem 0
lstLog.AddItem txt
SetListboxScrollbar
lstLog.ListIndex = lstLog.ListCount - 1
lstLog.Text = vbNullString
End Sub

Private Sub lstLog_DblClick()
On Error Resume Next
Clipboard.Clear
Clipboard.SetText lstLog.Text
End Sub

Private Function LoadFile1(strPath As String) As String
On Error GoTo E
If Dir$(strPath, vbHidden) = vbNullString Then Exit Function
With CreateObject("ADODB.Stream")
.Open
If ContainsUTF8(strPath) Then .Charset = "utf-8" Else: .Charset = "_autodetect_all"
.LoadFromFile strPath
LoadFile1 = .ReadText
End With
E:
End Function

Private Function LoadFile2(strPath As String) As Byte()
On Error GoTo E
Open strPath For Binary Access Read As #1
ReDim LoadFile2(LOF(1) - 1)
Get #1, , LoadFile2
Close #1
E:
End Function

Private Function ContainsUTF8(File As String) As Boolean
    On Error GoTo E

    Dim MLang       As CMultiLanguage
    Dim IMLang2     As IMultiLanguage2
    Dim Encoding()  As tagDetectEncodingInfo
    Dim encCount    As Long
    Dim inp()       As Byte
    Dim Index       As Long

    Open File For Binary As #1
    ReDim inp(LOF(1) - 1)
    Get #1, , inp()
    Close #1

    Set MLang = New CMultiLanguage
    Set IMLang2 = MLang

    encCount = 16
    ReDim Encoding(encCount - 1)
    IMLang2.DetectInputCodepage 0, 0, inp(0), UBound(inp) + 1, Encoding(0), encCount

    For Index = 0 To encCount - 1

        If Encoding(Index).nCodePage = 65001 Then 'UTF-8
            ContainsUTF8 = True
            Exit For
        End If

    Next
E:
    Set IMLang2 = Nothing
    Set MLang = Nothing
End Function

Public Function URLdecshort(ByRef Text As String) As String
On Error Resume Next
    Dim strArray() As String, lngA As Long
    strArray = Split(Replace(Text, "+", " "), "%")
    For lngA = 1 To UBound(strArray)
        strArray(lngA) = Chr$("&H" & Left$(strArray(lngA), 2)) & Mid$(strArray(lngA), 3)
    Next lngA
    URLdecshort = Join(strArray, vbNullString)
End Function

Private Function URLencshort(ByRef Text As String) As String
  Dim lngA As Long, strChar As String
  For lngA = 1 To Len(Text)
    strChar = Mid$(Text, lngA, 1)
    If Not strChar Like "[A-Za-z0-9]" Then
      strChar = "%" & Right$("0" & Hex$(Asc(strChar)), 2)
    ElseIf strChar = " " Then
      strChar = "+"
    End If
    URLencshort = URLencshort & strChar
  Next lngA
End Function

Function ProcessNumber(strT As Variant, Optional bolI As Boolean, Optional bolA As Boolean) As Integer
Dim intT As Integer
If Not bolI Then intT = 255 Else: intT = 32767
If strT <= intT Then
If strT < 0 Then ProcessNumber = strT * (-1) Else: ProcessNumber = strT
Else: If Not bolA Then ProcessNumber = intT Else: ProcessNumber = 0
End If
End Function

Private Function RegExpr(myPattern As String, myString As String, Optional myReplace As String, Optional bytResults As Byte, Optional intStart As Integer, Optional intCount As Integer) As Variant
'Modified by MikiSoft; Note: Must have "Microsoft VBScript Regular Expressions 5.5" library.
On Error GoTo E
Dim objRegExp As RegExp
Set objRegExp = New RegExp
objRegExp.Pattern = myPattern
objRegExp.IgnoreCase = True
objRegExp.Global = True
If bytResults > 0 Then
'If objRegExp.Test(myString) Then
Dim colMatches As MatchCollection
Set colMatches = objRegExp.Execute(myString)
If intCount = 0 Then intCount = colMatches.count - intStart
Dim i As Integer
For i = intStart To intStart + intCount - 1
If myReplace = vbNullString Then RegExpr = RegExpr & colMatches.Item(i).Value & vbNewLine Else: RegExpr = RegExpr & objRegExp.Replace(colMatches.Item(i).Value, myReplace) & vbNewLine
If i = colMatches.count - 1 Or bytResults = 2 Then Exit For
Next
RegExpr = Left$(RegExpr, Len(RegExpr) - 2)
'End If
Else: RegExpr = objRegExp.Test(myString)
End If
E: Set objRegExp = Nothing
End Function

Private Function RegExpr1(myPattern As String, myString As String, myReplace As String, bytResults As Byte, intStart As Integer, Optional intCount As Integer) As Variant
On Error GoTo E
Dim objRegExp As Object
Set objRegExp = CreateRG
objRegExp.Pattern = myPattern
objRegExp.IgnoreCase = True
If bytResults > 0 Then
'If objRegExp.Test(myString) Then
Dim strMatches() As String
strMatches = objRegExp.Execute(myString)
If UBound(strMatches) < LBound(strMatches) Then Exit Function
If intCount = 0 Then intCount = UBound(strMatches) + 1 - intStart
Dim i As Integer
For i = intStart To intStart + intCount - 1
If myReplace = vbNullString Then RegExpr1 = RegExpr1 & strMatches(i) & vbNewLine Else: RegExpr1 = RegExpr1 & objRegExp.Replace(strMatches(i), myReplace) & vbNewLine
If i = UBound(strMatches) Or bytResults = 2 Then Exit For
Next
RegExpr1 = Left$(RegExpr1, Len(RegExpr1) - 2)
'End If
Else: RegExpr1 = objRegExp.IsMatch(myString)
End If
E: Set objRegExp = Nothing
End Function

Private Function ReplaceString(ByVal strInp As String, Optional strSrc As String = vbNullChar) As String
If strInp = vbNullString Then Exit Function
strInp = Replace(strInp, "[nl]", vbNewLine)
'If InStr(strInp, "[inp") > 0 Then FindRI strInp
Dim strT As String, s() As String, intL As Integer, strR As String, strN As String 'strT(1) As String
'If bolR Then strT(0) = "rnd" Else: strT(0) = "inp"
s() = Split(strInp, "[rnd") '"[" & strT(0)
If UBound(s()) > 0 Then
Dim i As Byte
For i = 1 To UBound(s())
If InStr(s(i), "]") > 0 Then
strT = Left$(s(i), FindSep(s(i), , "]", "`") - 1)
If Len(strT) > 0 Then
intL = 0
AddChrs strR, strT, strInp, intL
If strR <> vbNullString Then strN = RandStr(strR, intL)
GoTo C
Else
strN = RandStr
C:
If strSrc = vbNullChar Then strN = Replace(strN, "'", "''")
strInp = Replace(strInp, "[rnd" & strT & "]", strN)
End If
End If
Next
End If
If InStr(strInp, "<") > 0 And InStr(strInp, ">") > 0 Then
Dim intC(1) As Long
intC(1) = 1
On Error GoTo N
Do
intC(0) = FindSep(strInp, intC(1), "<") + 1
If intC(0) = 1 Then Exit Do
intC(1) = FindSep(strInp, intC(0), ">")
If intC(1) = 0 Then Exit Do
strT = Mid$(strInp, intC(0), intC(1) - intC(0))
strN = vbNullString
On Error Resume Next
If Dir$(strT, vbHidden) <> vbNullString Then
strN = LoadFile1(strT)
If strSrc = vbNullChar Then strN = "'" & Replace(Replace(Replace(strN, "''", "'"), "'", "''"), "[src]", "['+'src]") & "'"
ElseIf strSrc <> vbNullChar Then GoTo N
End If
On Error GoTo 0
intC(0) = intC(0) - 1
strInp = Left$(strInp, intC(0) - 1) & Replace(strInp, "<" & strT & ">", strN, intC(0), 1)
intC(1) = intC(0) + Len(strN)
N:
Loop Until intC(1) > Len(strInp) - 3
End If
ReplaceString = Replace(strInp, "[dt]", Now)
If strSrc <> vbNullChar Then ReplaceString = Replace(ReplaceString, "[src]", strSrc)
'If InStr(strInp, "[rnd") > 0 Then FindRI ReplaceString ', True
End Function

Private Function AddChrs(strR As String, strT As String, Optional strInp As String, Optional intL As Integer, Optional intM As Variant) As String
Dim a As Byte, intT As Integer
On Error Resume Next
For a = 1 To Len(strT)
Select Case Mid$(strT, a, 1)
Case "D": strR = strR & strDigit
Case "L": strR = strR & strLett
Case "U": strR = strR & strULett
Case "M": strR = strR & strLett & strULett
Case "S": strR = strR & strSym
Case "`"
intT = FindC1(strT, a + 1, "`")
If intT = 0 Then Exit For
strR = strR & Replace(Mid$(strT, a + 1, intT - a - 1), "``", "`")
a = intT
Case Else
If Not IsMissing(intM) Then
Dim strT1 As String, s() As String
strT1 = Mid$(strT, a)
If IsNumeric(Replace(strT1, "-", vbNullString)) Then
If InStr(strT1, "-") > 0 Then
s() = Split(strT1, "-")
s(0) = Val(s(0))
s(1) = Val(s(1))
If CInt(s(0)) > CInt(s(1)) And CInt(s(1)) > 0 Then Exit For
If s(0) > 0 Then If intM(0) = 0 Then intM(0) = s(0) Else: If s(0) > intM(0) Then intM(0) = s(0)
If intM(1) = 0 Then intM(1) = s(1) Else: If s(1) < intM(1) Then intM(1) = s(1)
Else: If intM(1) = 0 Then intM(1) = strT1 Else: If strT1 < intM(1) Then intM(1) = strT1
End If
End If
ElseIf InStr(Mid$(strT, a), "-") > 0 Then strInp = Replace(strInp, "[rnd" & strT & "]", RandNum(Trim$(Split(strT, "-")(0)), Trim$(Split(strT, "-")(1))))
ElseIf IsNumeric(Mid$(strT, a)) Then intL = Mid$(strT, a)
End If
Exit For
End Select
Next
If Not IsMissing(intM) Then AddChrs = strR
End Function

Function ProcessString(ByVal strExp As String, strS As String, Optional intS As Integer = 1, Optional intStart As Integer, Optional intCount As Integer) As String
'On Error GoTo E
If strExp = vbNullString Then Exit Function
Dim intC(2) As Long, strT As String, intP As Long, bytC As Byte, strT2(2) As String, intL(1) As Long, tmpObj As IPluginInterface
strExp = strExp & "+"
R:
intC(0) = intS
Do
If InStr(intC(0), strExp, "+") > 0 Then
strT = Split(Mid$(strExp, intC(0)), "+")(0)
If strT = vbNullString Then
intC(0) = intC(0) + 1
GoTo N1
End If
End If
intC(1) = intC(0)
If Mid$(strExp, intC(0), 1) <> "'" And Not IsNumeric(strT) Then
If InStr(intC(0), strExp, "'") = 0 Then Exit Function
strExp = strExp & ")"
intS = intC(0)
Do
intC(2) = intC(1)
intC(0) = InStr(intC(2), strExp, "'") + 1
If intC(0) = 1 Then
intC(0) = Len(strExp) - 1
strT = Mid$(strExp, intS) & ")"
Else
intC(1) = FindC1(strExp, intC(0)) + 1
If intC(1) < 2 Then Exit Function
strT = Mid$(strExp, intC(2), intC(0) - intC(2) - 1)
End If
If InStr(strT, ")") > 0 Then
If intP = 0 Then Exit Function
If UBound(Split(strT, ")")) < bytC Then bytC = bytC - UBound(Split(strT, ")")) Else: bytC = 0
If InStr(intC(2), strExp, ")") < InStr(intC(2), strExp & "'", "'") Then intC(2) = InStr(intC(2), strExp, ")")
Dim strT1 As String
strT = Mid$(strExp, intP, intC(2) - intP + 1)
strT1 = strT
If Left$(strT, 3) = "rg(" Or Left$(strT, 4) = "rg1(" Then
If Left$(strT, 4) = "rg1(" Then intL(1) = 1 Else: intL(1) = 0
intL(0) = FindSep(strT, 4 + intL(1))
If intL(0) = 0 Then Exit Function
strT2(0) = ProcessString(Mid$(strT, 4 + intL(1), intL(0) - 4 - intL(1)), strS)
intL(0) = intL(0) + 1
intL(1) = FindSep(strT, intL(0))
If intL(1) = 0 Then intL(1) = Len(strT)
strT2(1) = ProcessString(Mid$(strT, intL(0), intL(1) - intL(0)), strS)
intL(0) = intL(1) + 1
strT2(2) = vbNullString
Dim intSt As Integer: intSt = 0
If intL(0) > Len(strT) - 1 Then GoTo N
If InStr("0123456789", Mid$(strT, intL(0), 1)) = 0 Then
intL(1) = FindSep(strT, intL(0))
If intL(1) = 0 Then intL(1) = Len(strT)
strT2(2) = ProcessString(Mid$(strT, intL(0), intL(1) - intL(0)), strS)
intL(0) = intL(1) + 1
If intL(0) > Len(strT) - 1 Then GoTo N
End If
intL(1) = FindSep(strT, intL(0))
If intL(1) > 0 Then
intSt = ProcessNumber(ProcessString(Mid$(strT, intL(0), intL(1) - intL(0)), strS), True, True)
intL(0) = intL(1) + 1
intL(1) = Len(strT)
Else: intSt = ProcessNumber(ProcessString(Mid$(Left$(strT, Len(strT) - 1), intL(0)), strS), True, True)
End If
N:
Dim intT As Integer: intT = 0 ': intT = intCount
If intStart = 0 Then
If intL(1) > intL(0) Then
'Dim intT1 As Integer
intT = ProcessNumber(ProcessString(Mid$(strT1, intL(0), intL(1) - intL(0)), strS), True, True) 'intT1
'If intT1 > intCount Then intCount = intT1
End If
If Left$(strT, 3) = "rg(" Then strT = RegExpr(strT2(1), strT2(0), strT2(2), 2, intSt, intT) Else: strT = RegExpr1(strT2(1), strT2(0), strT2(2), 2, intSt, intT)
If intCount < 1 Or intCount > intT Then intCount = intT
Else
If CLng(intSt) + CLng(intStart) < 32767 Then intT = intSt + intStart Else: intT = 32767
If Left$(strT, 3) = "rg(" Then strT = RegExpr(strT2(1), strT2(0), strT2(2), 2, intT) Else: strT = RegExpr1(strT2(1), strT2(0), strT2(2), 2, intT)
End If
strT = "'" & Replace(strT, "'", "''") & "'"
ElseIf Left$(strT, 4) = "rpl(" Then
intL(0) = FindSep(strT, 5)
If intL(0) = 0 Then Exit Function
strT2(0) = ProcessString(Mid$(strT, 5, intL(0) - 5), strS)
intL(0) = intL(0) + 1
intL(1) = FindSep(strT, intL(0))
If intL(1) = 0 Then Exit Function
strT2(1) = ProcessString(Mid$(strT, intL(0), intL(1) - intL(0)), strS)
intL(0) = intL(1) + 1
strT2(2) = ProcessString(Mid$(strT, intL(0), Len(strT) - intL(0)), strS)
strT = "'" & Replace(Replace(strT2(0), strT2(1), strT2(2)), "'", "''") & "'"
ElseIf InStr(strPlC, "," & Left$(strT, InStr(strT, "(") - 1) & ",") > 0 Then
ReDim strComm(0) As String
Dim i As Integer
strComm(0) = Left$(strT, InStr(strT, "(") - 1)
intL(0) = Len(strComm(0)) + 1
Do
intL(0) = intL(0) + 1
intL(1) = FindSep(strT, intL(0))
If intL(1) = 0 Then intL(1) = Len(strT)
ReDim Preserve strComm(UBound(strComm) + 1)
strComm(UBound(strComm)) = ProcessString(Mid$(strT, intL(0), intL(1) - intL(0)), strS)
intL(0) = intL(0) + (intL(1) - intL(0))
Loop Until intL(1) = Len(strT)
strT2(0) = Split(strPlC, "," & strComm(0) & ",")(0)
strT2(0) = Mid$(strT2(0), InStrRev(vbLf & strT2(0), vbLf, Len(strT2(0)) - 1) + 1)
Set tmpObj = Plugins.Item(Left$(strT2(0), InStr(strT2(0) & "|", "|") - 1))
bytPlgUse = bytPlgUse + 1
strT = "'" & Replace(tmpObj.Execute(strComm), "'", "''") & "'"
bytPlgUse = bytPlgUse - 1
If bolUnl And bytPlgUse = 0 Then Unload Me: Exit Function
Set tmpObj = Nothing
Else
strT2(0) = Mid$(strT, InStr(strT, "(") + 1)
strT2(0) = ProcessString(Left$(strT2(0), Len(strT2(0)) - 1), strS)
If Left$(strT, 5) = "dech(" Then
On Error GoTo E
Dim xml As Object
Set xml = CreateObject("MSXML2.DOMDocument.3.0")
With xml
.loadXML "<p>" & strT2(0) & "</p>"
strT = "'" & .selectSingleNode("p").nodeTypedValue & "'"
End With
Set xml = Nothing
E:
If err.Number <> 0 Then strT = "''"
On Error GoTo 0
ElseIf Left$(strT, 4) = "dec(" Then
strT = "'" & URLdecshort(strT2(0)) & "'"
ElseIf Left$(strT, 4) = "enc(" Then
strT = "'" & URLencshort(strT2(0)) & "'"
ElseIf Left$(strT, 2) = "u(" Then strT = "'" & Replace(UCase$(strT2(0)), "'", "''") & "'"
ElseIf Left$(strT, 2) = "l(" Then strT = "'" & Replace(LCase$(strT2(0)), "'", "''") & "'"
ElseIf Left$(strT, 4) = "b64(" Then strT = "'" & Encode64(strT2(0)) & "'"
ElseIf Left$(strT, 5) = "b64d(" Then strT = "'" & Decode64(strT2(0)) & "'"
ElseIf Left$(strT, 4) = "md5(" Then strT = "'" & LCase$(Hash.HashBytes(StrConv(strT2(0), vbFromUnicode))) & "'"
ElseIf Left$(strT, 4) = "num(" Then strT = Val(strT2(0))
End If
'Else
'E: strT = "''"
End If
If strT1 = strT Then Exit Function
strExp = Replace(strExp, strT1, strT)
bytC = 0
GoTo R
ElseIf InStr(strT, "(") > 0 Then
bytC = bytC + UBound(Split(strT, "("))
intP = intC(0) - Len(Mid$(strT, InStrRev(strT, "("))) - 1
Do
intP = intP - 1
If intP = 0 Then Exit Do
Loop Until Mid$(strExp, intP, 1) = "," Or Mid$(strExp, intP, 1) = "+" Or Mid$(strExp, intP, 1) = "("
intP = intP + 1
End If
Loop Until bytC = 0
intC(1) = intC(1) - 1
ElseIf Mid$(strExp, intC(0), 1) = "'" Then
intC(0) = intC(0) + 1
intC(1) = FindC1(strExp, intC(0))
If intC(1) < 2 Then Exit Function
ProcessString = ProcessString & Replace(Replace(Mid$(strExp, intC(0), intC(1) - intC(0)), "''", "'"), "[src]", strS)
Else
intC(1) = InStr(intC(0), strExp, "+")
If intC(1) > 0 Then
If IsNumeric(ProcessString) Then ProcessString = ProcessString + Val(Mid$(strExp, intC(0), intC(1) - intC(0))) Else: ProcessString = ProcessString & Mid$(strExp, intC(0), intC(1) - intC(0))
intC(1) = intC(1) - 1
Else: Exit Do
End If
End If
intC(0) = intC(1) + 2
'If IsNumeric(ProcessString) And IsNumeric(strT) Then ProcessString = CDbl(ProcessString + strT) Else: ProcessString = ProcessString & strT
N1:
Loop Until intC(0) >= Len(strExp)
'E:
End Function

Private Function FindC1(strS As String, Optional intC As Long = 2, Optional strC As String = "'") As Long
FindC1 = InStr(intC, Left$(strS, intC - 1) & Replace(Mid$(strS, intC), strC & strC, "  "), strC)
End Function

Private Function FindC(ByVal strS As String, Optional ByVal intC As Integer = 2) As Integer
FindC = InStr(intC, Left$(strS, intC - 1) & Replace(Mid$(strS, intC), strC, "  "), """")
End Function

Private Function FindSep(strExp As String, Optional intS As Long = 1, Optional strC As String = ",", Optional strE As String = "'") As Long
Dim intC(2) As Long
intC(1) = intS
intC(2) = intS
Do
intC(0) = intC(2)
intC(1) = InStr(intC(2), strExp, strE) + 1
If intC(1) = 1 Then Exit Do
intC(2) = FindC1(strExp, intC(1), strE) + 1
Loop Until InStr(Mid$(strExp, intC(0), intC(1) - intC(0) - 1), strC) > 0 Or intC(2) = 1
If intC(2) = 1 Then FindSep = InStr(intC(1), strExp, strC) Else: FindSep = InStr(intC(0), strExp, strC)
End Function

Private Function RandNum(ByVal Low As Long, ByVal High As Long) As Long
RandNum = Int((High - Low + 1) * Rnd) + Low
End Function

Private Function RandStr(Optional strR As String, Optional intL As Integer) As String
    If strR = vbNullString Then strR = strDigit & strLett & strULett & strSym
    If intL = 0 Then intL = 15
    Dim i As Integer
    For i = 1 To intL
        RandStr = RandStr & Mid$(strR, Int(Rnd() * Len(strR) + 1), 1)
    Next
End Function

Private Function CreateRG() As Object
Dim ActCtx As ACTCTX_
Dim res As Boolean, actHandle As Long, actCookie As Long
ActCtx.cbSize = Len(ActCtx)
ActCtx.dwFlags = ACTCTX_FLAG_RESOURCE_NAME_VALID
ActCtx.lpSource = App.path & IIf(Right$(App.path, 1) <> "\", "\", vbNullString) & App.EXEName & ".exe" & vbNullChar
ActCtx.lpResourceName = "#3"
actHandle = CreateActCtx(ActCtx)
If actHandle = INVALID_HANDLE_VALUE Then Exit Function
res = ActivateActCtx(actHandle, actCookie)
If Not res Then
Call ReleaseActCtx(actHandle)
actHandle = 0
Exit Function
End If
On Error Resume Next
Set CreateRG = CreateObject("DotNetCOMRegExLib.DotNetRegEx")
Call ReleaseActCtx(actHandle)
Call DeactivateActCtx(0, actCookie)
actHandle = 0
actCookie = 0
End Function

Private Function TrimComma(strT As String) As String
TrimComma = strT
Do While InStr(TrimComma, ",,") > 0
TrimComma = Replace(TrimComma, ",,", ",")
Loop
TrimComma = Replace(Replace(Replace(Replace(TrimComma, "(", vbNullString), ")", vbNullString), "+", vbNullString), "'", vbNullString)
End Function

Private Function get_relative_path_to(ByVal child_path As String, Optional folder As Boolean, Optional parent_path As String) As String

If parent_path = vbNullString Then parent_path = strInitD

If LCase$(Left$(child_path, 1)) <> LCase$(Left$(parent_path, 1)) Or InStr(child_path, "\") = 0 Then
get_relative_path_to = child_path
Exit Function
End If

Dim attr As Long
If folder Then attr = FILE_ATTRIBUTE_DIRECTORY Else: attr = FILE_ATTRIBUTE_NORMAL
Dim out_str As String
Dim par_str As String
Dim child_str As String

out_str = String$(MAX_PATH, 0)

par_str = parent_path + String$(100, 0)
child_str = child_path + String$(100, 0)

PathRelativePathTo out_str, par_str, FILE_ATTRIBUTE_DIRECTORY, child_str, attr

out_str = StripTerminator(out_str)

If Left$(out_str, 2) <> ".\" Then
If folder Then If Right$(out_str, 1) <> "\" Then out_str = out_str & "\"
If UBound(Split(out_str, "..\")) = UBound(Split(parent_path & IIf(Right$(parent_path, 1) <> "\", "\", vbNullString), "\")) - 1 Then out_str = Mid$(out_str, InStrRev(out_str, "..\") + 2) Else: If Left$(out_str, 1) = "\" Then out_str = Mid$(out_str, 2)
Else: out_str = Mid$(out_str, 3)
End If
If Len(out_str) > 1 Then If Right$(out_str, 1) = "\" Then out_str = Left$(out_str, Len(out_str) - 1)

get_relative_path_to = out_str
End Function

'Remove all trailing Chr$(0)'s
Private Function StripTerminator(sInput As String) As String
    Dim ZeroPos As Long
    ZeroPos = InStr(1, sInput, Chr$(0))
    If ZeroPos > 0 Then
        StripTerminator = Left$(sInput, ZeroPos - 1)
    Else
        StripTerminator = sInput
    End If
End Function

Private Sub SetListboxScrollbar()
Dim new_len As Long
Static max_len As Long
If lstLog.ListCount > 0 Then

        new_len = 10 + lstLog.Parent.ScaleX( _
            lstLog.Parent.TextWidth(lstLog.list(lstLog.ListCount - 1)), _
            lstLog.Parent.ScaleMode, vbPixels)
        If max_len < new_len Then
        max_len = new_len
E:
        SendMessage lstLog.hwnd, _
        LB_SETHORIZONTALEXTENT, _
        max_len, 0
        End If
Else
max_len = 0
GoTo E
End If
End Sub

' Purpose:  Take a string whose bytes are in the byte array <the_abytCPString>, with code page <the_nCodePage>, convert to a VB string.
Private Function FromCPString(ByRef the_abytCPString() As Byte, ByVal the_nCodePage As Long) As String

    Dim sOutput                     As String
    Dim nValueLen                   As Long
    Dim nOutputCharLen              As Long

    ' If the code page says this is already compatible with the VB string, then just copy it into the string. No messing.
    If the_nCodePage = CP_UTF16_LE Then
        FromCPString = the_abytCPString()
    Else

        ' Cache the input length.
        nValueLen = UBound(the_abytCPString) - LBound(the_abytCPString) + 1

        ' See how big the output buffer will be.
        nOutputCharLen = MultiByteToWideChar(the_nCodePage, 0&, VarPtr(the_abytCPString(LBound(the_abytCPString))), nValueLen, 0&, 0&)

        ' Resize output byte array to the size of the UTF-8 string.
        sOutput = Space$(nOutputCharLen)

        ' Make this API call again, this time giving a pointer to the output byte array.
        MultiByteToWideChar the_nCodePage, 0&, VarPtr(the_abytCPString(LBound(the_abytCPString))), nValueLen, StrPtr(sOutput), nOutputCharLen

        ' Return the array.
        FromCPString = sOutput

    End If

End Function

' Purpose:  Converts a VB string (UTF-16) to UTF8 - as a binary array.
Private Function ToCPString(ByRef the_sValue As String, Optional ByVal the_nCodePage As Long = CP_ACP) As Byte()

    Dim abytOutput()                As Byte
    Dim nValueLen                   As Long
    Dim nOutputByteLen              As Long

    If the_nCodePage = CP_UTF16_LE Then
        ToCPString = the_sValue
    Else

        ' Cache the input length.
        nValueLen = Len(the_sValue)

        ' See how big the output buffer will be.
        nOutputByteLen = WideCharToMultiByte(the_nCodePage, 0&, StrPtr(the_sValue), nValueLen, 0&, 0&, 0&, 0&)

        If nOutputByteLen > 0 Then
            ' Resize output byte array to the size of the UTF-8 string.
            ReDim abytOutput(1 To nOutputByteLen)

            ' Make this API call again, this time giving a pointer to the output byte array.
            WideCharToMultiByte the_nCodePage, 0&, StrPtr(the_sValue), nValueLen, VarPtr(abytOutput(1)), nOutputByteLen, 0&, 0&
        End If

        ' Return the array.
        ToCPString = abytOutput()

    End If

End Function

Private Sub CatBinary(bytData() As Byte, Bytes() As Byte)
    Dim BytesLen As Long, BinaryNext As Long
    
    BinaryNext = UBound(bytData) + 1
    BytesLen = UBound(Bytes) - LBound(Bytes) + 1
    If BinaryNext + BytesLen > BinaryNext Then ReDim Preserve bytData(BinaryNext + BytesLen - 1)
    CopyMemory VarPtr(bytData(BinaryNext)), VarPtr(Bytes(LBound(Bytes))), BytesLen
End Sub
 
Private Sub CatBinaryString(bytData() As Byte, Text As String)
    Dim Bytes() As Byte
    
    Bytes = ToCPString(Text)
    CatBinary bytData, Bytes
End Sub

Private Function CopyPointerToString(ByVal in_pString As Long) As String

    Dim nLen            As Long

    ' Need to copy the data at the string pointer to a VB string buffer.
    ' Get the length of the string, allocate space, and copy to that buffer.

    nLen = lstrlen(in_pString)
    CopyPointerToString = Space$(nLen)
    CopyMemory StrPtr(CopyPointerToString), in_pString, nLen * 2

End Function

Private Function GetMimeTypeFromData(ByRef in_abytData() As Byte, ByRef in_sProposedMimeType As String) As String

    Dim nLBound          As Long
    Dim nUBound          As Long
    Dim pMimeTypeOut     As Long
    Dim nRet             As Long

    nLBound = LBound(in_abytData)
    nUBound = UBound(in_abytData)

    nRet = FindMimeFromData(0&, 0&, VarPtr(in_abytData(nLBound)), nUBound - nLBound + 1, StrPtr(in_sProposedMimeType), FMFD_DEFAULT, pMimeTypeOut, 0&)

    If nRet = S_OK Then
        GetMimeTypeFromData = CopyPointerToString(pMimeTypeOut)
        CoTaskMemFree pMimeTypeOut
    Else
        GetMimeTypeFromData = "application/octet-stream"
    End If

End Function

Private Sub SystemTray_MouseUp(Button As Integer)
 SetForegroundWindow Me.hwnd
 If Button = 1 Then
  On Error GoTo E
  Me.Visible = True
  SystemTray.RemoveIcon
  cmdMintoTray.Checked = False
 'Else: PopupMenu mnu1
 End If
Exit Sub
E: MsgBox "Can't do that right now!", vbExclamation
End Sub

Private Sub cmdMintoTray_Click()
Me.Visible = False
If App.LogMode > 0 Then
SystemTray.Tip = Me.Caption
SystemTray.AddIcon
End If
cmdMintoTray.Checked = True
End Sub

Private Sub chkOnTop_Click()
Static tm As Boolean
tm = Not tm
  If tm Then
  chkOnTop.Checked = True
    SetTopMostWindow Me.hwnd, True
    Me.Caption = Me.Caption & " [on top]"
  Else
  chkOnTop.Checked = False
    SetTopMostWindow Me.hwnd, False
    Me.Caption = Left$(Me.Caption, InStrRev(Me.Caption, " [on top]") - 1) & Mid$(Me.Caption, InStrRev(Me.Caption, " [on top]") + 9)
  End If
End Sub

Private Sub Form_Load()
    Dim hIcon As Long, hInst As Long

    If App.LogMode Then
        Set Icon = Nothing
        hInst = App.hInstance
        hWndOwner = GetWindow(hwnd, GW_OWNER)

        hIcon = LoadImageA(hInst, 1&, IMAGE_ICON, , , LR_DEFAULTSIZE Or LR_SHARED)
If App.LogMode > 0 Then
Set SystemTray = New clsInTray
SystemTray.hIcon = hIcon
strInitD = CurDir$ & IIf(Right$(CurDir$, 1) <> "\", "\", vbNullString)
Else
cmdMintoTray.Enabled = False
SetCurrentDirectoryA App.path
strInitD = App.path & IIf(Right$(App.path, 1) <> "\", "\", vbNullString)
End If
        DestroyIcon SendMessageA(hwnd, WM_SETICON, ICON_BIG, hIcon)
        DestroyIcon SendMessageA(hWndOwner, WM_SETICON, ICON_BIG, hIcon)

        hIcon = LoadImageA(hInst, 1&, IMAGE_ICON, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON), LR_SHARED)
        DestroyIcon SendMessageA(hwnd, WM_SETICON, ICON_SMALL, hIcon)
        DestroyIcon SendMessageA(hWndOwner, WM_SETICON, ICON_SMALL, hIcon)
    End If
Dim intC(2) As Integer
Rnd -Timer * Now
Randomize
strC = String$(2, """")
Set rh = New cAsyncRequests
bytTOrigin0 = 5
bytThreads = 1
bytSubThr = 1
bytDelay = 1
bytTimeout = 20
bytTOrigin0 = 5
bytTOrigin1 = 5
intLogMax = 32767
Dim strT As String, s1() As String, s() As String, cnt As Byte, strT1 As String, i As Byte, a As Byte, tmpObj As IPluginInterface
cnt = 103
s1() = Split(StrConv(LoadResData(101, 10), vbUnicode), vbNewLine)
On Error GoTo N
For i = 0 To UBound(s1)
If s1(i) = vbNullString Then GoTo N
s1(i) = Trim$(s1(i))
If InStr(";#[", Left$(s1(i), 1)) = 0 Then
If strT1 = vbNullString Then
strT = Mid$(s1(i), InStr(s1(i), "=") + 1)
If strT <> vbNullString Then
s1(i) = LCase$(Left$(s1(i), InStr(s1(i), "=") - 1))
If IsNumeric(strT) Then
If strT < 0 Then strT = strT * (-1)
Select Case s1(i)
Case "limit": bytLimit = ProcessNumber(strT - 1)
Case "silent": bytSilent = ProcessNumber(strT)
Case "meltortray": bolMT = strT > 0
Case "skiploadingerrors": bolSkipErr = strT > 0
Case "pluginsdir": bolTmp = strT > 0
Case "threads": bytThreads = ProcessNumber(strT)
Case "subthreads": bytSubThr = ProcessNumber(strT)
Case "timeout": bytTimeout = ProcessNumber(strT)
Case "donotretry": bolNoRetry = strT > 0
Case "delaybetweenretries": bytDelay = ProcessNumber(strT)
Case "maxretriespercycle": bytMaxR = ProcessNumber(strT)
Case "after": intAfter = ProcessNumber(strT, True)
Case "debug": bolDebug = strT > 0
End Select
ElseIf s1(i) = "executebatch" Then
strCmd = strT
ElseIf s1(i) = "savelog" Then
strPath(0) = strT
ElseIf s1(i) = "saveoutput" Then
strPath(1) = strT
ElseIf s1(i) = "title" Then
App.Title = strT
Me.Caption = strT
ElseIf s1(i) = "after" Then
intAfter = Left$(strT, Len(strT) - 1)
bolHours = True
ElseIf s1(i) = "originmax" Then
If Right$(strT, 1) = "c" Then
bolColl = True
strT = Left$(strT, Len(strT) - 1)
End If
If strT <> vbNullString Then
If Right$(strT, 1) = "n" Then
bolNoEach = True
strT = Left$(strT, Len(strT) - 1)
End If
If strT <> vbNullString Then
strT = Trim$(strT)
If Left$(strT, 1) <> ";" Then bytTOrigin0 = ProcessNumber(Left$(strT, InStr(strT & ";", ";") - 1))
If InStr(strT, ";") > 0 Then
bytTOrigin1 = ProcessNumber(Mid$(strT, InStr(strT, ";") + 1))
If bytTOrigin1 = 0 Then bolColl = True
End If
End If
End If
ElseIf s1(i) = "logoutputmax" Then
strT = Trim$(strT)
If Left$(strT, 1) <> ";" Then intLogMax = ProcessNumber(Left$(strT, InStr(strT & ";", ";") - 1), True)
If InStr(strT, ";") > 0 Then intOutMax = ProcessNumber(Mid$(strT, InStr(strT, ";") + 1), True)
ElseIf s1(i) = "output" Then
intC(1) = -1
For a = 0 To 1
intC(0) = intC(1) + 3
intC(1) = FindC(strT, intC(0))
If intC(1) = 0 Then Exit For
If a = 0 Then strTemplate0 = Replace(Replace(Mid$(strT, intC(0), intC(1) - intC(0)), "[nl]", vbNewLine), strC, """") Else: strTemplate1 = Replace(Replace(Mid$(strT, intC(0), intC(1) - intC(0)), "[nl]", vbNewLine), strC, """")
Next
ElseIf s1(i) = "loadedplugin" Then
s() = Split(strT, "|")
If UBound(s) > 0 Then
strT1 = DropPlugin(s(0), cnt)
If strT1 = vbNullString Then If ChkErr(True) Then Exit Sub Else: GoTo N
On Error GoTo -1
On Error GoTo N1
For a = 1 To UBound(s)
Set tmpObj = CreateObjectEx2(strT1, strT1, s(a))
If Not tmpObj Is Nothing Then
Plugins.Add tmpObj, s(0) & "/" & s(a)
strPlC = strPlC & "|" & s(0) & "/" & s(a) & "|," & TrimComma(tmpObj.Startup(Me)) & "," & vbLf
cnt = cnt + 1
Set tmpObj = Nothing
ElseIf ChkErr(True) Then Exit Sub
End If
Next
ElseIf DropPlugin(s(0), cnt, True) <> vbNullString Then
bolRg = True
cnt = cnt + 1
ElseIf ChkErr(True) Then Exit Sub
End If
N1:
strT1 = vbNullString
If err.Description <> vbNullString Then If ChkErr(True) Then Exit Sub Else: err.Clear
On Error GoTo -1
On Error GoTo N
End If
End If
Else: Print #2, s1(i)
End If
ElseIf Left(s1(i), 1) = "[" Then
s1(i) = Mid$(s1(i), 2, Len(s1(i)) - 2)
a = InStr(s1(i), "/")
If a > 1 And a < Len(s1(i)) Then
If strT1 <> vbNullString Then GoSub ExecP
strT = s1(i)
strT1 = Left$(strT, a - 1)
strT1 = Left$(strT, InStr(strT & ".", ".") - 1) & ".ini"
If bolTmp Then strT1 = Environ$("TMP") & "\" & strT1
On Error Resume Next
'SetAttr strT1, vbNormal
'Kill strT1
Open strT1 For Output Access Write As #2
If err.Description <> vbNullString Then If ChkErr Then Exit Sub Else: err.Clear
On Error GoTo N
End If
End If
N:
Next
On Error GoTo 0
Dim strL As String
If strT1 <> vbNullString Then
ExecP:
Close #2
strL = DropPlugin(Left$(strT, InStr(strT, "/") - 1), cnt)
strDrP = strDrP & strT1 & vbLf
If strL = vbNullString Then
If False Then
N3:
On Error GoTo -1
On Error GoTo N
End If
If Not ChkErr(True) Then
If Not bolTmp Then SetAttr strT1, vbNormal
Kill strT1
GoTo C
Else: Exit Sub
End If
End If
If Not bolTmp Then SetAttr strT1, vbHidden
On Error GoTo -1
On Error GoTo N3
Set tmpObj = CreateObjectEx2(strL, strL, Mid$(strT, InStr(strT, "/") + 1))
If Not tmpObj Is Nothing Then
Plugins.Add tmpObj, strT
strPlC = strPlC & "|" & strT & "|," & TrimComma(tmpObj.Startup(Me)) & "," & vbLf
cnt = cnt + 1
Set tmpObj = Nothing
ElseIf ChkErr(True) Then
If Not bolTmp Then SetAttr strT1, vbNormal
Kill strT1
Exit Sub
End If
C:
On Error Resume Next
Return
On Error GoTo 0
End If
If InStr(Command$, """") > 0 Then
intC(1) = 1
intC(2) = 1
Do
R:
intC(0) = intC(2)
intC(1) = InStr(intC(2), Command$, """")
If intC(1) = intC(2) Then
intC(2) = intC(2) + 1
GoTo R
End If
If intC(1) = 0 Then intC(1) = Len(Command$) + 1
If intC(0) = intC(1) Then Exit Do
DetectC intC(0), intC(1)
intC(2) = FindC(Command$, intC(1) + 1) + 1
If intC(2) = 1 Then Exit Do
Loop Until InStr(Mid$(Command$, intC(0), intC(1) - intC(0) - 1), """") > 0
ElseIf Command$ <> vbNullString Then DetectC 1, Len(Command$) + 1
End If
If bytSilent > 0 Then
App.TaskVisible = False
bolLO(0) = strPath(0) = vbNullString
bolLO(1) = strPath(1) = vbNullString
End If
If bolRg Then
Dim objT As Object
Set objT = CreateRG
If Not objT Is Nothing Then
On Error GoTo N2
objT.Pattern = "."
objT.Execute "."
Set objT = Nothing
Open App.path & IIf(Right$(App.path, 1) <> "\", "\", vbNullString) & "dotnetcomregexlib.dll" For Input Access Read As #2
strPlC = ",rg1," & vbLf & strPlC
N2:
If err.Description <> vbNullString Then If ChkErr Then Exit Sub Else: err.Clear
On Error GoTo 0
End If
End If
If CInt(bytThreads) + CInt(bytSubThr) > 255 Then If bytSubThr > bytThreads Then bytSubThr = bytSubThr - bytThreads Else: bytSubThr = 1
DimP True
If bytSilent = 0 Then If Not bolMT Then Me.Visible = True Else: cmdMintoTray_Click
Screen.MousePointer = 11
lblStatus.Caption = "Loading configuration..."
lblStatus.Refresh
Dim strI As String, k As Integer, j As Integer, b As Integer, bytT As Byte
strI = "|"
s1() = Split(StrConv(LoadResData(102, 10), vbUnicode), vbNewLine)
On Error Resume Next
For k = 0 To UBound(s1)
If s1(k) = vbNullString Then GoTo Nx1
s1(k) = Trim$(Replace(s1(k), vbCr, vbNullString))
If s1(k) <> vbNullString Then
inp:
If InStr(";#[", Left$(s1(k), 1)) = 0 Then
If InStr(s1(k), "=") > 0 Then
strL = Left$(s1(k), InStr(s1(k), "=") - 1)
If InStr(strI, "|" & strL & "|") = 0 Then
s1(k) = Mid$(Left$(s1(k), Len(s1(k))), Len(strL) + 2)
Select Case strL
Case "url"
If InStr(strI, "|strings|") > 0 Then
If InStr(strPO, "-" & bytIC & "-") > 0 Then
Ad:
strI = "|"
bytIC = bytIC + 1
If bytIC > bytLimit Then AddL
End If
ElseIf InStr(strI, "|if|") > 0 Then GoTo Ad
End If
strURLData(bytIC) = s1(k)
If Not ChkURL(strURLData(bytIC)) Then strI = strI & "url|"
Case "post"
If s1(k) <> vbNullString Then
strURLData(bytIC) = strURLData(bytIC) & vbLf & s1(k)
strI = strI & "post|"
End If
Case "if"
If InStr(strI, "|strings|") > 0 Then
If InStr(strPO, "-" & bytIC & "-") > 0 Then
strI = "|"
bytIC = bytIC + 1
If bytIC > bytLimit Then AddL
End If
End If
If Right$(s1(k), 1) <> """" Then s1(k) = s1(k) & """"
intC(0) = 2
intC(1) = FindC(s1(k), intC(0))
j = 0
strT = vbNullString
Do While Len(s1(k)) > intC(1)
If j > 0 Then
strT = Mid$(s1(k), intC(0), 1)
If Not IsNumeric(strT) Then strT = 0 Else: If strT < 0 Then strT = 0 Else: If strT > 1 Then strT = 1
strT = vbLf & strT
intC(0) = intC(0) + 3
intC(1) = FindC(s1(k), intC(0))
If intC(1) = 0 Then Exit Do
If j > (bytLimit + 1) \ 2 Then AddL
End If
strIf(bytIC, j) = Trim$(Replace(Mid$(s1(k), intC(0), intC(1) - intC(0)), strC, """"))
intC(0) = intC(1) + 2
strT1 = Mid$(s1(k), intC(0), 1)
If Not IsNumeric(strT1) Then strT1 = 0 Else: If strT1 < 0 Or strT1 > 5 Then strT1 = 0
strIf(bytIC, j) = strIf(bytIC, j) & vbLf & strT1
intC(0) = intC(0) + 3
intC(1) = FindC(s1(k), intC(0))
If intC(1) = 0 Then strIf(bytIC, j) = vbNullString: Exit Do
strT1 = Trim$(Replace(Mid$(s1(k), intC(0), intC(1) - intC(0)), strC, """"))
If strIf(bytIC, j) = vbNullString And strT1 = vbNullString Then
strIf(bytIC, j) = vbNullString
bytSh(bytIC) = j + 1
If j > 0 Or strT > 0 Then GoTo Ni
Else: strIf(bytIC, j) = strIf(bytIC, j) & vbLf & strT1 & strT
End If
j = j + 1
Ni:
intC(0) = intC(1) + 2
Loop
If j = bytSh(bytIC) Then bytSh(bytIC) = 0
strI = strI & "if|"
Case "strings"
If Right$(s1(k), 1) <> """" Then s1(k) = s1(k) & """"
intC(0) = 2
intC(1) = FindC(s1(k), intC(0))
j = 0
Dim bolS As Boolean
Do While Len(s1(k)) > intC(1)
If Mid$(s1(k), intC(0) - 1, 1) <> """" Then
intC(1) = intC(0) + 7
If Mid$(s1(k), intC(1), 1) <> """" Then Exit Do
strT = vbLf & Mid$(s1(k), intC(0) - 1, intC(1) - intC(0))
s() = Split(strT, ",")
If UBound(s) = 3 Then
If s(1) = "1" Or s(3) = "1" Then bolS = True
Else: strT = vbLf
End If
intC(0) = intC(1) + 1
Else: strT = vbLf
End If
intC(1) = FindC(s1(k), intC(0))
If intC(1) = 0 Then Exit Do
strStrings(bytIC, j) = Trim$(Replace(Replace(Replace(Replace(Replace(Mid$(s1(k), intC(0), intC(1) - intC(0)), strC, """"), "%", ""), "{", ""), "}", ""), "'", ""))
If strStrings(bytIC, j) <> vbNullString Then
For b = 0 To j - 1
If Split(strStrings(bytIC, b), vbLf)(0) = strStrings(bytIC, j) Then strStrings(bytIC, j) = vbNullString: GoTo Ns1
Next
Else
Ns1:
intC(0) = FindC(s1(k), FindC(s1(k), intC(1) + 3)) + 1
GoTo ns
End If
intC(0) = intC(1) + 3
intC(1) = FindC(s1(k), intC(0))
If intC(1) = 0 Then strStrings(bytIC, j) = vbNullString: Exit Do
strT1 = Trim$(Replace(Mid$(s1(k), intC(0), intC(1) - intC(0)), strC, """"))
intC(0) = intC(1) + 1
If strT1 <> vbNullString Then
strStrings(bytIC, j) = strStrings(bytIC, j) & vbLf & strT1 & strT
If bolS Then
Dim intT As Integer
If InStr(strPO, "-" & bytIC & "-") > 0 Then
intT = Split(Split(strPO, "-" & bytIC & "-")(1), vbLf)(0)
strPO = Replace(strPO, "-" & bytIC & "-" & intT, "-" & bytIC & "-" & intT + 1, , 1)
Else: strPO = strPO & "-" & bytIC & "-1" & vbLf
End If
bolS = False
End If
j = j + 1
If j > (bytLimit + 1) \ 2 Then AddL
Else: strStrings(bytIC, j) = vbNullString
End If
ns:
If Mid$(s1(k), intC(0), 1) <> ";" Then Exit Do
intC(0) = intC(0) + 2
Loop
strI = strI & "strings|"
Case "headers"
If Right$(s1(k), 1) <> """" Then s1(k) = s1(k) & """"
intC(0) = 2
intC(1) = FindC(s1(k), intC(0))
j = 0
Do While Len(s1(k)) > intC(1)
intC(1) = FindC(s1(k), intC(0))
strHeaders(bytIC, j) = Trim$(Replace(Mid$(s1(k), intC(0), intC(1) - intC(0)), strC, """"))
If strHeaders(bytIC, j) <> vbNullString Then
For b = 0 To j - 1
If Split(strHeaders(bytIC, b), vbLf)(0) = strHeaders(bytIC, j) Then strHeaders(bytIC, j) = vbNullString: GoTo Nh1
Next
Else
Nh1:
intC(0) = FindC(s1(k), FindC(s1(k), intC(1) + 3)) + 2
GoTo NH
End If
intC(0) = intC(1) + 3
intC(1) = FindC(s1(k), intC(0))
If intC(1) = 0 Then strHeaders(bytIC, j) = vbNullString: Exit Do
strT1 = Trim$(Replace(Mid$(s1(k), intC(0), intC(1) - intC(0)), strC, """"))
intC(0) = intC(1) + 2
If strT1 <> vbNullString Then
strHeaders(bytIC, j) = strHeaders(bytIC, j) & vbLf & strT1
j = j + 1
If j > (bytLimit + 1) \ 2 Then AddL
Else: strHeaders(bytIC, j) = vbNullString
End If
NH:
If Mid$(s1(k), intC(0), 1) <> """" Then Exit Do
intC(0) = intC(0) + 1
Loop
strI = strI & "headers|"
Case "wait"
If Left$(s1(k), 1) <> """" Then
If Split(s1(k), ";")(0) = vbNullString Then
If s1(k) <> vbNullString Then strWait(0, bytIC) = Val(s1(k))
GoTo Nx
Else: strWait(0, bytIC) = Val(Split(s1(k), ";")(0))
End If
bytT = Len(strWait(0, bytIC)) + 2
If bytT > Len(s1(k)) Then GoTo Nx
Else
strWait(0, bytIC) = Mid$(s1(k), 2, FindC(s1(k)) - 2)
bytT = Len(strWait(0, bytIC)) + 4
If bytT > Len(s1(k)) Then GoTo Nx
End If
If Mid$(s1(k), bytT, 1) = """" Then
strWait(1, bytIC) = Mid$(s1(k), bytT + 1, FindC(s1(k), bytT + 1) - bytT - 1)
Else: strWait(1, bytIC) = Val(Mid$(s1(k), bytT))
End If
Nx:
If strWait(0, bytIC) = "0" Then strWait(0, bytIC) = vbNullString
If strWait(1, bytIC) = "0" Then strWait(1, bytIC) = vbNullString
strI = strI & "wait|"
Case "goto"
s() = Split(s1(k) & ";", ";")
If IsNumeric(s(0)) Then intGoto(0, bytIC) = Abs(CInt(s(0)))
If IsNumeric(s(1)) Then intGoto(1, bytIC) = Abs(CInt(s(1)))
strI = strI & "goto|"
End Select
Else
Cl:
ChkAd strI, k
GoTo inp
End If
End If
ElseIf Left$(s1(k), 1) = "[" Then ChkAd strI, k
End If
End If
Nx1:
Next
On Error GoTo 0
If Not Filled(bytIC) Then bytIC = bytIC - 1
Screen.MousePointer = 0
lblStatus.Caption = "Idle..."
If Not bolLO(0) Then addLog "Program started."
cmdStart_Click
End Sub

Private Sub ChkAd(strI As String, i As Integer)
strI = "|"
If Not Filled(bytIC) Then
Dim a As Byte
Do While strHeaders(i, a) <> vbNullString
strHeaders(i, a) = vbNullString
a = a + 1
If UBound(strHeaders, 2) < a Then Exit Do
Loop
a = 0
Do While strStrings(i, a) <> vbNullString
strStrings(i, a) = vbNullString
a = a + 1
If UBound(strStrings, 2) < a Then Exit Do
Loop
For a = 0 To UBound(strIf, 2)
If bytSh(i) < a Then Exit For
strIf(i, a) = vbNullString
Next
For a = 0 To 1
strWait(a, i) = vbNullString
intGoto(a, i) = 0
Next
If bolDebug And Not bolLO(0) Then addLog "Index " & i + 1 & " removed.", True
Else
bytIC = bytIC + 1
If bytIC > bytLimit Then AddL
End If
End Sub

Private Function ChkErr(Optional bolT As Boolean) As Boolean
If bolSkipErr Then Exit Function
If bytSilent < 2 Then MsgBox IIf(bolT, "There was error while loading plugins! Check if firewall/anti-virus software is blocking this program.", "There was error while loading .NET RegEx plugin! Check if .NET Framework 3.5 is properly installed," & vbNewLine & "or if firewall/anti-virus software is blocking this program."), vbCritical
Unload Me
ChkErr = True
End Function

Private Function DropPlugin(strN As String, bytR As Byte, Optional bolT As Boolean) As String
On Error GoTo E
Dim bytF() As Byte: bytF = LoadResData(bytR, 10)
If bolT Then DropPlugin = App.path & IIf(Right$(App.path, 1) <> "\", "\", vbNullString) & strN Else: If Not bolTmp Then DropPlugin = strN Else: DropPlugin = Environ("TMP") & "\" & strN
If Dir$(DropPlugin, vbHidden) = vbNullString Then
Open DropPlugin For Binary Access Write As #1
Put #1, , bytF
Close #1
If Not bolTmp Or bolT Then SetAttr DropPlugin, vbHidden
End If
If Not bolT Then strDrP = strDrP & DropPlugin & vbLf
Exit Function
E: DropPlugin = vbNullString
End Function

Private Sub DetectC(intS As Integer, intE As Integer)
Dim strT(1) As String
strT(0) = " " & Mid$(Command$, intS, intE - intS) & " "
Do
If InStr(1, strT(0), " -k ", vbTextCompare) > 0 Then
bolSkipErr = True
strT(0) = Replace(strT(0), " -k", vbNullString)
ElseIf InStr(1, strT(0), " -m ", vbTextCompare) > 0 Then
bolMT = True
strT(0) = Replace(strT(0), " -m", vbNullString)
ElseIf InStr(1, strT(0), " -d ", vbTextCompare) > 0 Then
bolDebug = True
strT(0) = Replace(strT(0), " -d", vbNullString)
ElseIf InStr(1, strT(0), " -e ", vbTextCompare) > 0 Then strCmd = ExtrF("e", intS, strT(0))
ElseIf InStr(1, strT(0), " -o ", vbTextCompare) > 0 Then
strT(1) = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(ExtrF("o", intS, strT(0)), "\", vbNullString), "/", vbNullString), ":", vbNullString), "*", vbNullString), "?", vbNullString), "<", vbNullString), ">", vbNullString), "|", vbNullString)
If strT(1) <> vbNullString Then strPath(1) = strT(1) Else: strPath(1) = "results.txt"
ElseIf InStr(1, strT(0), " -l ", vbTextCompare) > 0 Then
strT(1) = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(ExtrF("l", intS, strT(0)), "\", vbNullString), "/", vbNullString), ":", vbNullString), "*", vbNullString), "?", vbNullString), "<", vbNullString), ">", vbNullString), "|", vbNullString)
If strT(1) <> vbNullString Then strPath(0) = strT(1) Else: strPath(0) = "{NOW}"
ElseIf InStr(1, strT(0), " -t ", vbTextCompare) > 0 Then AddVal "t", strT(0)
ElseIf InStr(1, strT(0), " -h ", vbTextCompare) > 0 Then AddVal "h", strT(0)
ElseIf InStr(1, strT(0), " -u ", vbTextCompare) > 0 Then AddVal "u", strT(0)
ElseIf InStr(1, strT(0), " -a ", vbTextCompare) > 0 Then AddVal "a", strT(0)
ElseIf InStr(1, strT(0), " -r ", vbTextCompare) > 0 Then AddVal "r", strT(0)
ElseIf InStr(1, strT(0), " -i ", vbTextCompare) > 0 Then AddVal "i", strT(0)
ElseIf InStr(1, strT(0), " -n ", vbTextCompare) > 0 Then
bolNoRetry = True
strT(0) = Replace(strT(0), " -n", vbNullString)
ElseIf InStr(1, strT(0), " -f ", vbTextCompare) > 0 Or InStr(1, strT(0), " -fh ", vbTextCompare) > 0 Then
strT(1) = Split(Split(strT(0), " -")(1), " ")(0)
bolHours = Right$(strT(1), 1) = "h"
Dim strT1 As String: strT1 = Split(Mid$(strT(0), Len(strT(1)) + 5), " ")(0)
intAfter = ProcessNumber(strT1, True)
strT(0) = Replace(strT(0), " -" & strT(1) & " " & strT1, vbNullString)
Else: GoTo E
End If
Loop Until Trim$(strT(0)) = vbNullString
Exit Sub
E:
End Sub

Private Sub AddVal(strCh As String, strT1 As String)
Dim strT As String, bytT As Byte
strT = Split(Split(strT1, " -" & strCh & " ", , vbTextCompare)(1), " ")(0)
If IsNumeric(strT) Then
bytT = ProcessNumber(strT)
Select Case strCh
Case "t": bytTimeout = bytT
Case "h": bytThreads = bytT
Case "u": bytSubThr = bytT
Case "a": bytDelay = bytT
Case "r": bytMaxR = bytT
Case "i": bytSilent = bytT
End Select
End If
strT1 = Replace(strT1, " -" & strCh & " " & strT, vbNullString, , , vbTextCompare)
End Sub

Private Function ExtrF(strT As String, intS As Integer, strT1 As String) As String
strT1 = Replace(strT1, " -" & strT & " ", vbNullString, , , vbTextCompare)
Dim intC(1) As Byte
intC(0) = InStr(intS, " " & Command$, " -" & strT & " " & """", vbTextCompare) + 4
If intC(0) = 4 Then Exit Function
intC(1) = FindC(Command$, intC(0) + 1)
If intC(1) = 0 Then Exit Function
ExtrF = Replace(Mid$(Command$, intC(0), intC(1) - intC(0)), strC, """")
End Function

Private Function ChkURL(ByVal strT As String, Optional bolT As Boolean) As Boolean
If strT <> vbNullString Then
'If InStr(strT, vbLf) = 0 And InStr(strT, vbCr) = 0 Then
If Left$(strT, 1) = "%" And InStr(2, strT, "%") > 0 Or Left$(strT, 4) = "[inp" And InStr(2, strT, "]") > 0 Then Exit Function
If ChkStr(strT) Then bolT = True Else: Exit Function
If Left$(strT, 7) = "http://" Or Left$(strT, 8) = "https://" Then Exit Function
'If Left$(strT, 7) = "http://" Then
'If Len(strT) >= 11 Then Exit Function
'ElseIf Len(strT) >= 12 Then Exit Function
'End If
'End If
End If
'End If
ChkURL = True
End Function

Private Function Filled(bytIndex As Byte) As Boolean
If strURLData(bytIndex) <> vbNullString Then Filled = Not ChkURL(Split(strURLData(bytIndex), vbLf)(0))
If Filled Then Exit Function
If InStr(strIf(bytIndex, 0), vbLf) > 0 Then
If Split(strIf(bytIndex, 0), vbLf)(0) <> vbNullString Or Split(strIf(bytIndex, 0), vbLf)(2) <> vbNullString Then Filled = True
If Filled Then Exit Function
End If
If InStr(strPO, "-" & bytIndex & "-") > 0 Then Filled = True
End Function

Private Function AddL() As Boolean
Dim intL As Integer
If bytLimit + 3 <= 256 Then intL = bytLimit + 3 Else: Exit Function
Dim strHeaders1() As String
Dim strStrings1() As String
Dim strIf1() As String
ReDim strHeaders1(bytLimit, (bytLimit + 1) \ 2)
ReDim strStrings1(bytLimit, (bytLimit + 1) \ 2)
ReDim strIf1(bytLimit, (bytLimit + 1) \ 2)
Dim i As Byte, j As Byte
For i = 0 To bytLimit
For j = 0 To (bytLimit + 1) \ 2
strHeaders1(i, j) = strHeaders(i, j)
strStrings1(i, j) = strStrings(i, j)
strIf1(i, j) = strIf(i, j)
Next
Next
Dim bytL As Byte: bytL = bytLimit
bytLimit = intL - 1
DimP
For i = 0 To bytL
For j = 0 To bytL \ 2
strHeaders(i, j) = strHeaders1(i, j)
strStrings(i, j) = strStrings1(i, j)
strIf(i, j) = strIf1(i, j)
Next
Next
Erase strHeaders1
Erase strStrings1
Erase strIf1
If bolDebug And Not bolLO(0) Then addLog "Limit moved, from " & bytL + 1 & " to " & intL & " (added " & intL - (bytL + 1) & " free spaces).", True
AddL = True
End Function

Private Sub DimP(Optional bolN As Boolean)
If Not bolN Then
ReDim Preserve bytSh(bytLimit)
ReDim Preserve strURLData(bytLimit)
ReDim Preserve strWait(1, bytLimit)
ReDim Preserve intGoto(1, bytLimit)
Else
ReDim bytSh(bytLimit)
ReDim strURLData(bytLimit)
ReDim strWait(1, bytLimit)
ReDim intGoto(1, bytLimit)
End If
ReDim strHeaders(bytLimit, (bytLimit + 1) \ 2)
ReDim strStrings(bytLimit, (bytLimit + 1) \ 2)
ReDim strIf(bytLimit, (bytLimit + 1) \ 2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Enabled = False
bolUnl = True
If bytPlgUse > 0 Then
lblStatus.Caption = "Waiting for plugins..."
Cancel = 1
Exit Sub
End If
If Not bolEx Then
If lblStatus.Caption <> "Stopping..." Then
bolAb = True
rh.Cleanup
Else
bolEx = True
Exit Sub
End If
End If
lblStatus.Caption = "Exiting..."
Set rh = Nothing
If bytSilent = 0 And Not Me.Visible And App.LogMode > 0 Then SystemTray.RemoveIcon
Set SystemTray = Nothing
Set Plugins = Nothing
Dim s() As String, i As Integer
s() = Split(strDrP, vbLf)
If bolRg Then Close #2
On Error Resume Next
Dim strE As String, bytN As Byte, lngI As Long
For i = 0 To UBound(s) - 1
If InStr(strE, """" & s(i) & """") = 0 Then
If Right$(s(i), 4) <> ".ini" Then
If Not UnloadLibrary(s(i)) Then
strE = strE & ":rp" & bytN & vbCrLf & "del /a """ & s(i) & """" & vbCrLf & "if not exist """ & s(i) & """ set p" & bytN & "=1" & vbCrLf & "if not defined p" & bytN & " goto rp" & bytN & vbCrLf
bytN = bytN + 1
lngI = InStrRev(s(i), ".")
If lngI > InStrRev(s(i), "\") Then s(i) = Left$(s(i), lngI)
s(i) = Left$(s(i), InStrRev(s(i), ".")) & "ini"
strE = strE & ":rp" & bytN & vbCrLf & "del /a """ & s(i) & """" & vbCrLf & "if not exist """ & s(i) & """ set p" & bytN & "=1" & vbCrLf & "if not defined p" & bytN & " goto rp" & bytN & vbCrLf
bytN = bytN + 1
End If
End If
If Not bolTmp Then SetAttr s(i), vbNormal
Kill s(i)
End If
Next
On Error GoTo E
If (Not bolMT Or bytSilent = 0) And Not bolRg And strE = vbNullString Then Exit Sub
If bolTmp Then SetCurrentDirectoryA Environ("tmp") Else: SetCurrentDirectoryA App.path
Open "DelMe.bat" For Output Access Write As #1
If bolMT Then
Print #1, ":r"
Print #1, "del """ & App.EXEName & ".exe" & """"
Print #1, "if not exist """ & App.EXEName & ".exe" & """ set d=1"
Print #1, "if not defined d goto r"
End If
If bolRg Then
Print #1, ":r1"
Print #1, "del /a:h dotnetcomregexlib.dll"
Print #1, "if not exist dotnetcomregexlib.dll set d1=1"
Print #1, "if not defined d1 goto r1"
End If
If strE <> vbNullString Then Print #1, strE
Print #1, "del /a:h %0"
Close #1
SetAttr "DelMe.bat", vbHidden
Shell "DelMe.bat", vbHide
E:
End Sub

Private Sub disF()
If Not cmdStart.Enabled Then
If lstLog.list(0) <> vbNullString Then
cmdSave(0).Enabled = True
cmdClear(0).Enabled = True
End If
If txtOutput.Text <> vbNullString Then
cmdSave(1).Enabled = True
cmdClear(1).Enabled = True
End If
cmdStart.Enabled = True
Else
Dim i As Byte
For i = 0 To 1
cmdSave(i).Enabled = False
cmdClear(i).Enabled = False
Next
cmdStart.Enabled = False
End If
End Sub

Private Sub cmdStart_Click()
If cmdStart.Caption = "Start" Then
lblStatus.Caption = "Preparing..."
lblStatus.Refresh
Screen.MousePointer = 11
disF
Me.Caption = Me.Caption & " (working)"
If cmdMintoTray.Checked And App.LogMode > 0 Then SystemTray.Tip = Me.Caption
rh.Timeout = bytTimeout * 1000
Dim bytT As Byte, i As Integer
bytT = bytThreads - 1
intSubT = bytSubThr
If intSubT = 0 Then intSubT = 255
Set colSrc = New Collection
Set colPubStr = New Collection
Set colStr = New Collection
Set colMax = New Collection
Set colMaxR = New Collection
Set colInput = New Collection
Dim b As Byte ', strT As String
For i = 0 To bytIC
If PrepareInput("URL" & vbLf & strURLData(i), bytT, i) Then GoTo E
If InStr(strURLData(i), vbLf) > 0 Then If PrepareInput("Post" & vbLf & Split(strURLData(i), vbLf)(1), bytT, i) Then GoTo E
b = 0
Do While strHeaders(i, b) <> vbNullString
If PrepareInput("Header name" & vbLf & strHeaders(i, b), bytT, i, b) Then GoTo E
If PrepareInput("Header value" & vbLf & Split(strHeaders(i, b), vbLf)(1), bytT, i, b) Then GoTo E
b = b + 1
If UBound(strHeaders, 2) < b Then Exit Do
Loop
b = 0
Do While strStrings(i, b) <> vbNullString
If PrepareInput(strStrings(i, b), bytT, i, b, True) Then GoTo E
b = b + 1
If UBound(strStrings, 2) < b Then Exit Do
Loop
b = 0
'strT = vbNullString
Dim s() As String
Do While strIf(i, b) <> vbNullString
s() = Split(strIf(i, b), vbLf)
'If b > 0 Then strT = cmbOper(0).List(s(3)) & " "
If PrepareInput("If A (" & b & ")" & vbLf & s(0), bytT, i, b) Then GoTo E
If PrepareInput("If A <=> [B] (" & b & ")" & vbLf & s(2), bytT, i, b) Then GoTo E 'strT & "If A " & cmbSign(0).List(s(1)) & " [B]" & vbLf & s(2)
b = b + 1
If UBound(strIf, 2) < b Or bytSh(i) < b Then Exit Do
Loop
If PrepareInput("Then/Else wait seconds" & vbLf & strWait(0, i), bytT, i, 0) Then GoTo E
If PrepareInput("Then/Else wait seconds" & vbLf & strWait(1, i), bytT, i, 1) Then GoTo E
Next
Dim tmr As Object
If strURLData(0) <> vbNullString Then Set tmr = tmrU Else: Set tmr = tmrI: tmrI(0).Tag = vbNullString
bytOrigin = bytTOrigin1
If bolNoEach Then bytOrigin = bytOrigin \ (bytT + 1)
If intAfter > 0 Then datCompl = DateAdd(IIf(bolHours, "h", "n"), intAfter, Now)
bolAb = False
bytActive = 0
If CurDir$ <> strInitD Then
strLastPath = CurDir$
SetCurrentDirectoryA strInitD
End If
For i = 0 To bytT
If i > 0 Then Load tmr(i)
intTmrCount = intTmrCount + 1
tmr(i).Enabled = True
Next
Set tmr = Nothing
cmdStart.Caption = "Stop"
cmdStart.Enabled = True
If Not bolLO(0) Then addLog "Process started."
lblStatus.Caption = "..."
Screen.MousePointer = 0
ElseIf Not bolAb Then
cmdStart.Enabled = False
lblStatus.Caption = "Stopping..."
Enb
End If
Exit Sub
E: Enb 2
End Sub

Private Function PrepareInput(strI As String, bytT As Byte, i As Integer, Optional b As Byte, Optional bolT As Boolean) As Boolean
Dim strT(3) As String, intM(1) As Integer, t As Byte
strT(0) = Split(strI, vbLf)(1)
If InStr(strT(0), "[inp") = 0 Then Exit Function
If InStr(Split(strT(0), "[inp")(1), "]") = 0 Then Exit Function
If bolT Then strT(3) = IIf(Val(Split(Split(strI, vbLf)(2) & ",", ",")(0)) = "1", 2, 0) + Val(Split(Split(strI, vbLf)(2) & ",,,", ",")(2)) & "%" & Split(strI, vbLf)(0) & "%" Else: strT(3) = "0" & Split(strI, vbLf)(0)
If PrepareCol(colInput, "-1," & strT(3)) <> vbNullString Then Exit Function
R:
frmInput.strInf = strT(3) & vbLf & t + 1 & "," & i + 1 & "," & b + 1 & vbLf
strT(2) = strT(0)
If strT(1) = vbNullString Then
Do
strT(1) = Split(strT(0), "[inp")(1)
strT(1) = Left$(strT(1), FindSep(strT(1), , "]", "`") - 1)
If Len(strT(1)) > 0 Then
frmInput.strInf = AddChrs(frmInput.strInf, strT(1), , , intM)
strT(2) = Replace(strT(2), "[inp" & strT(1) & "]", "[inp]")
End If
strT(0) = Replace(strT(0), "[inp" & strT(1) & "]", vbNullString)
Loop Until InStr(strT(0), "[inp") = 0 Or InStr(strT(0), "]") = 0
End If
If intM(0) > 0 Or intM(1) > 0 Then
frmInput.strInf = frmInput.strInf & vbLf
If intM(0) > 0 Then frmInput.strInf = frmInput.strInf & intM(0) & "-"
If intM(1) > 0 Then frmInput.strInf = frmInput.strInf & intM(1)
End If
Screen.MousePointer = 0
lblStatus.Caption = "Waiting for user input."
frmInput.Show vbModal
Screen.MousePointer = 11
If bolT Then If Val(Split(Split(strI, vbLf)(2) & ",", ",")(0)) = "1" Or Split(strI, vbLf)(0) = "URL" Then If Mid$(frmInput.strInf, 2) = vbNullString Then If Replace(strT(2), "[inp]", vbNullString) = vbNullString Then PrepareInput = True: Exit Function
lblStatus.Caption = "Preparing..."
lblStatus.Refresh
If Left$(frmInput.strInf, 1) = "1" Then strT(0) = t & "," Else: strT(0) = "-1,"
colInput.Add Left$(LTrim$(strT(3)), 1) & strT(2) & vbLf & Mid$(frmInput.strInf, 2), strT(0) & i & Mid$(LTrim$(strT(3)), 2)
If Left$(frmInput.strInf, 1) = "0" Then Exit Function
If Left$(strT(3), 1) <> " " Then strT(3) = " " & strT(3)
t = t + 1
If t > bytT Then Exit Function
strT(0) = Split(strI, vbLf)(1)
GoTo R
End Function

Private Sub Enb(Optional R As Byte)
bolAb = True
Do While intTmrCount > 0
DoEvents
Loop
If R = 0 Then
bytActive = 255
rh.Cleanup
End If
If bolEx Then
Unload Me
Exit Sub
End If
Set colSrc = Nothing
Set colPubStr = Nothing
Set colStr = Nothing
Set colMax = Nothing
Set colMaxR = Nothing
Set colInput = Nothing
Set colCurrO = Nothing
intLTmr(0) = 0
intLTmr(1) = 0
If strLastPath <> vbNullString Then
SetCurrentDirectoryA strLastPath
strLastPath = vbNullString
End If
Dim strT As String
Select Case R
Case 0: If Not cmdStart.Enabled Then strT = "aborted" Else: strT = "stopped"
Case 1: strT = "finished"
Case 2: strT = "canceled"
Case 3: strT = "automatically aborted"
End Select
If Not bolLO(0) Then addLog "Process " & strT & "."
lblStatus.Caption = "Idle..."
Screen.MousePointer = 0
Me.Caption = Left$(Me.Caption, InStrRev(Me.Caption, " (working)") - 1) & Mid$(Me.Caption, InStrRev(Me.Caption, " (working)") + 10)
If cmdMintoTray.Checked And App.LogMode > 0 Then SystemTray.Tip = Me.Caption
cmdStart.Caption = "Start"
cmdStart.Enabled = False
disF
If Me.Visible Then cmdStart.SetFocus
If R = 2 Then Exit Sub
If strPath(0) <> vbNullString Then cmdSave_Click 0
If strPath(1) <> vbNullString Then cmdSave_Click 1
If strCmd <> vbNullString Then
'If r = 0 Then If MsgBox("Proceed with executing Batch commands?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
lblStatus.Caption = "Executing Batch commands..."
Shell "cmd.exe /c " & Replace(Replace(strCmd, "%TLocation%", App.path & IIf(Right$(App.path, 1) <> "\", "\", vbNullString)), "%TFilename%", App.EXEName & ".exe"), IIf(bytSilent = 0, vbMaximizedFocus, vbHidden)
If Not bolLO(0) Then addLog "Batch commands executed."
lblStatus.Caption = "Idle..."
ElseIf (R = 1 Or R = 3) And Not bolDebug And Me.Visible Then MsgBox "The process has finished successfully!", vbInformation
End If
If Not Me.Visible Or Me.WindowState = vbMinimized Then Unload Me
End Sub

Private Sub SubmitReq(a As Byte, j As Byte, Optional O As String, Optional i As Integer, Optional strT1 As String)
Dim strU(1) As String, strD(1) As String, varD As Variant, bolD1 As Boolean, bolT As Boolean, strS As String, strH(1) As String, colH As Collection, Key As String, strTrim As String
If bytOrigin > 0 Then
On Error Resume Next
colCurrO.Add "", O
On Error GoTo 0
End If
GetSrc strS, CStr(j), O, a
strU(0) = Split(strURLData(a), vbLf)(0)
If bolAb Then Exit Sub
Dim strT2 As String: If Left$(O, 1) = "/" Or InStr(strT1, "-") > 0 Or InStr(strT1, "+") > 0 Then strT2 = strT1
Dim strCurr As String: strCurr = "{T: " & j + 1 & ", S: " & i & ", I: " & a + 1 & ", O:" & strT2 & "} "
Dim intT As Integer
If i = 0 And Left$(O, 1) <> "/" Then intT = 255 Else: intT = intSubT
If intT - bytActive < 1 Then GoTo E1
strTrim = TrimO(O & " " & i)
strU(1) = ProceedString(strU(0), strS, a, j, i, O, strT2, strTrim, -2)
If strU(1) = vbNullString Then
If i = 0 Then
If Not bolLO(0) Then addLog "{T: " & j + 1 & ", I: " & a + 1 & ", O:" & strT2 & "} Error: Blank URL!"
GoTo E
Else: GoTo E3
End If
ElseIf strU(0) <> strU(1) Then
If ChkURL(strU(1)) Then
If Not bolLO(0) Then addLog strCurr & "Error: Invalid or blank URL!", True
GoTo E
End If
End If
strD(0) = Split(strURLData(a) & vbLf, vbLf)(1)
If strD(0) <> vbNullString Then bolD1 = True
If intT - bytActive > 0 Then
Dim b As Byte
Do
If bolDebug And Not bolLO(0) Then If strU(1) <> strU(0) Then addLog strCurr & "URL: " & strU(1), True
b = 0
Set colH = New Collection
Do While strHeaders(a, b) <> vbNullString
strH(0) = Replace(ProceedString(Split(strHeaders(a, b), vbLf)(0), strS, a, j, i, O, strT2, strTrim, -4), vbLf, vbNullString)
If StrPtr(strH(0)) = 0 Then GoTo E
strH(1) = Replace(ProceedString(Split(strHeaders(a, b), vbLf)(1), strS, a, j, i, O, strT2, strTrim, -5), vbLf, vbNullString)
If StrPtr(strH(1)) = 0 Then GoTo E
If strH(0) <> vbNullString And strH(1) <> vbNullString Then
If intT - bytActive < 1 Then GoTo E1
colH.Add strH(0) & vbLf & strH(1)
If bolDebug And Not bolLO(0) Then If strH(0) <> Split(strHeaders(a, b), vbLf)(0) Or strH(1) <> Split(strHeaders(a, b), vbLf)(1) Then addLog strCurr & strH(0) & ": " & strH(1), True
ElseIf bolDebug And Not bolLO(0) Then addLog strCurr & "Warning: Unexpected blank add. header at " & b + 1 & ".", True
End If
b = b + 1
If UBound(strHeaders, 2) < b Then Exit Do
Loop
If intT - bytActive < 1 Then Exit Do
If bolD1 Then
If Left$(strD(0), 1) = "[" And Right$(strD(0), 1) = "]" And InStr(strD(0), ":") > 0 Then
Dim strC As String, intC(1) As Long, strBoundary As String, bytD() As Byte, bolB As Boolean, strT(1) As String
If intT - bytActive < 1 Then GoTo E1
strBoundary = "--" & RandStr(strDigit & strLett & strULett)
intC(0) = 2
Do While intC(0) < Len(strD(0))
bolT = Not bolT
If bolT Then strC = ":" Else: strC = ";"
intC(1) = FindSep(strD(0), intC(0), strC)
If intT - bytActive < 1 Then GoTo E1
If intC(1) = 0 Then intC(1) = Len(strD(0))
strT(0) = Mid$(strD(0), intC(0), intC(1) - intC(0))
If Not bolT Then
If Left$(strT(0), 1) = "<" And Right$(strT(0), 1) = ">" Then
strT(0) = Mid$(strT(0), 2, Len(strT(0)) - 2)
If Not bolB Then
If intT - bytActive < 1 Then GoTo E1
bytD = ""
CatBinaryString bytD, strD(1)
strD(1) = vbNullString
bolB = True
End If
If intT - bytActive < 1 Then GoTo E1
Dim bytF() As Byte: bytF = ""
On Error GoTo N1
If strT(0) <> vbNullString Then
If Dir$(strT(0), vbHidden) <> vbNullString Then
bytF = LoadFile2(strT(0))
CatBinaryString bytD, "; filename=""" & Mid$(strT(0), InStrRev(strT(0), "\") + 1) & """" & vbCrLf & "Content-Type: " & GetMimeTypeFromData(bytF, vbNullString) & vbCrLf & vbCrLf
If intT - bytActive < 1 Then GoTo E1
CatBinary bytD, bytF
Erase bytF
CatBinaryString bytD, vbCrLf
Else
On Error GoTo 0
N1:
If bolDebug And Not bolLO(0) Then addLog "{T: " & j + 1 & ", I: " & a + 1 & ", O:" & strT2 & "} Warning: File can't be opened or doesn't exist: " & strT(0), True
CatBinaryString bytD, vbCrLf & vbCrLf & vbCrLf
End If
Else: CatBinaryString bytD, "; filename=""""" & vbCrLf & "Content-Type: application/octet-stream" & vbCrLf & vbCrLf
End If
Else
strT(1) = vbCrLf & vbCrLf & ProceedString(strT(0), strS, a, j, i, O, strT2, strTrim, -3) & vbCrLf
If Not bolB Then strD(1) = strD(1) & strT(1) Else: CatBinaryString bytD, strT(1)
End If
Else
If strT(1) <> vbNullString Then strT(1) = vbCrLf
strT(1) = strBoundary & vbCrLf & "Content-Disposition: form-data; name=""" & ProceedString(strT(0), strS, a, j, i, O, strT2, strTrim, -3) & """"
If Not bolB Then strD(1) = strD(1) & strT(1) Else: CatBinaryString bytD, strT(1)
End If
If intT - bytActive < 1 Then GoTo E1
intC(0) = intC(1) + 1
Loop
If Not bolB Then
varD = strD(1) & strBoundary & "--" & vbCrLf
strD(1) = vbNullString
Else
CatBinaryString bytD, strBoundary & "--" & vbCrLf
varD = bytD
Erase bytD
bolB = False
End If
If VarType(varD) = vbEmpty Then
If bolDebug And Not bolLO(0) Then If i > 0 Then If VarPtr(varD) <> 0 Then addLog strCurr & "Warning: Unexpected end of pipe (POST).", True
GoTo E
End If
If intT - bytActive < 1 Then Exit Do
If bolDebug And Not bolLO(0) Then addLog strCurr & "POST: (multipart/form-data)", True
colH.Add "Content-Type" & vbLf & "multipart/form-data; boundary=" & Mid$(strBoundary, 3)
ElseIf Left$(strD(0), 1) = "<" And Right$(strD(0), 1) = ">" Then
strD(0) = Mid$(strD(0), 2, Len(strD(0)) - 2)
On Error GoTo N
If Dir$(strD(0), vbHidden) = vbNullString Then
N:
If Not bolLO(0) Then addLog "{T: " & j + 1 & ", I: " & a + 1 & ", O:" & strT2 & "} Error: File can't be opened or doesn't exist: " & strD(0), True
GoTo E2
Else: varD = LoadFile2(strD(0))
End If
On Error GoTo 0
If intT - bytActive < 1 Then GoTo E1
If bolDebug And Not bolLO(0) Then addLog strCurr & "PUT: " & strD(0), True
strU(1) = "*" & strU(1)
Else
If strD(0) <> "''" Then varD = ProceedString(strD(0), strS, a, j, i, O, strT2, strTrim, -3) Else: varD = vbNullChar
If intT - bytActive < 1 Then Exit Do
If bolDebug And Not bolLO(0) Then If strD(0) <> varD Then addLog strCurr & "POST: " & Replace(varD, vbLf, "[nl]"), True
If Left$(strD(0), 1) = "{" And Right$(strD(0), 1) = "}" Then colH.Add "Content-Type" & vbLf & "application/json" Else: colH.Add "Content-Type" & vbLf & "application/x-www-form-urlencoded"
End If
End If
'Dim strT3 As String: If strT1 = vbNullString Then strT3 = IIf(strT2 <> vbNullString Or i > 0, o, vbNullString)
Key = j & "," & i & "," & a & "," & O & "," & strT1 & "," & strTrim
If intT - bytActive < 1 Then Exit Do
On Error GoTo err
rh.AddRequest(Key).SendRequest strU(1), varD, colH
On Error GoTo 0
bytActive = bytActive + 1
If tmrQ.Enabled Then tmrQ.Tag = Replace(tmrQ.Tag, "-" & a & "," & j & "," & O & "," & i & "," & strT1 & vbLf, vbNullString, , 1)
If Not bolLO(0) Then addLog strCurr & "Request sent."
bytD = ""
Set colH = Nothing
i = i + 1
strCurr = Replace(strCurr, "S: " & i - 1 & ",", "S: " & i & ",")
If i > Val(PrepareCol(colMax, a & "," & j & "," & O)) Then Exit Sub
'Debug.Print i, Val(PrepareCol(colMax, a & "," & J & "," & o)), J
If i = 1 Then intT = intSubT
If intT - bytActive < 1 Then Exit Do
strTrim = TrimO(O & " " & i)
strU(1) = ProceedString(strU(0), strS, a, j, i, O, strT2, strTrim, -2)
If strU(1) = vbNullString Then
E3:
If bolDebug And Not bolLO(0) Then If StrPtr(strU(1)) <> 0 Then addLog strCurr & "Warning: Unexpected end of pipe (URL).", True
GoTo E
End If
Loop Until intT - bytActive < 1
End If
E1:
If Not bolAb And InStr(tmrQ.Tag, "-" & a & "," & j & "," & O & "," & i & "," & strT1 & vbLf) = 0 Then
'If Not strU(1) = vbNullString Then 'And Not bolD1 Or bolD1 And Not strU(1) = vbNullString And Not varD = vbNullString
tmrQ.Tag = tmrQ.Tag & "-" & a & "," & j & "," & O & "," & i & "," & strT1 & vbLf
If bolDebug And Not bolLO(0) Then addLog strCurr & "Added to queue.", True
If Not tmrQ.Enabled Then tmrQ.Enabled = True
'End If
End If
GoTo E2
E: If tmrQ.Enabled Then If InStr(tmrQ.Tag, "-" & a & "," & j & "," & O & "," & i & "," & strT1 & vbLf) > 0 Then tmrQ.Tag = Replace(tmrQ.Tag, "-" & a & "," & j & "," & O & "," & i & "," & strT1 & vbLf, vbNullString, , 1): Exit Sub
E2:
PrepTr CStr(j), CStr(i), CStr(a), O, True
If bolUnl Then Exit Sub
If Not tmrQ.Enabled Then ChkT Else: ChkSt
Exit Sub
err:
If err.Number <> 457 Then
rh.RemoveRequest Key
If Not bolLO(0) Then addLog strCurr & "Error: " & err.Description & ".", True
GoTo E
Else: GoTo E1
End If
End Sub

Private Sub GetSrc(strS As String, j As String, ByVal O As String, a As Byte)
If bytOrigin > 0 Then
Dim strT(2) As String, i As Integer, intT(1) As Integer
O = Replace(Replace(O, "|", " "), "/", vbNullString, , 1)
For i = colSrc.count To 1 Step -1
strT(1) = colSrc.Item(i)
If Left$(strT(1), Len(j) + 1) = j & "," Then
strT(1) = Mid$(strT(1), Len(j) + 2)
strT(1) = Left$(strT(1), InStr(strT(1), "\") - 1)
strT(2) = Replace(Replace(strT(1), "|", " "), "/", vbNullString, , 1)
If InStr(O & " ", strT(2) & " ") = 1 Then intT(0) = Len(O) - Len(strT(2)) Else: If InStr(" " & StrReverse(O), " " & StrReverse(Replace(strT(2), "...", vbNullString, , 1))) = 1 Then intT(0) = Len(O) - Len(Replace(strT(2), "...", vbNullString, , 1)) Else: If InStr(O, "...") > 0 Or InStr(strT(2), "...") > 0 Then intT(0) = -Len(strT(1)) Else: intT(0) = -1
If intT(0) <> -1 Then
If strT(0) <> vbNullString Then
Select Case True
Case intT(0) = 0, intT(0) < intT(1): GoTo G
End Select
Else
G:
strT(0) = strT(1)
strS = Mid$(colSrc.Item(i), Len(strT(1)) + Len(j) + 3)
If intT(0) = 0 Then Exit For Else: intT(1) = intT(0)
End If
End If
End If
Next
Else
On Error GoTo E
Do
strS = PrepareCol(colSrc, j & "," & a & "," & O)
If a > 0 And strS = vbNullString Then strS = PrepareCol(colSrc, j & "," & a - 1 & "," & O)
If strS = vbNullString Then O = Left$(O, InStrRev(O, " ") - 1) Else: Exit Do
Loop Until O = vbNullString
E:
End If
End Sub

Private Sub rh_ResponseFinished(Req As cAsyncRequest)
If bolAb Then Exit Sub
Dim s() As String, strS As String
strS = Req.http.Status & " " & Req.http.StatusText & vbNewLine & Req.http.GetAllResponseHeaders
On Error Resume Next
strS = strS & FromCPString(Req.http.ResponseBody, CP_UTF8)
On Error GoTo 0
s() = Split(Req.Key, ",")
rh.RemoveRequest Req
bytActive = bytActive - 1
If bolDebug Then
Dim strT As String: If Left$(s(3), 1) = "/" Or InStr(s(4), "-") > 0 Or InStr(s(4), "+") > 0 Then strT = s(4)
If Not bolLO(0) Then addLog "{T: " & s(0) + 1 & ", S: " & s(1) & ", I: " & s(2) + 1 & ", O:" & strT & "} Response status - " & Split(strS, vbNewLine)(0), True
End If
CheckIf s(0), s(1), s(2), s(3), strS, s(4), s(5)
End Sub

Private Sub rh_Error(Req As cAsyncRequest, ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
rh.RemoveRequest Req
bytActive = bytActive - 1
If bolAb Then Exit Sub
Dim s() As String: s() = Split(Req.Key, ",")
Dim strT(1) As String
If Left$(s(3), 1) = "/" Or InStr(s(4), "-") > 0 Or InStr(s(4), "+") > 0 Then strT(1) = s(4)
strT(0) = "T: " & s(0) + 1 & ", S: " & s(1) & ", I: " & s(2) + 1 & ", O:" & strT(1)
If Not bolLO(0) Then addLog "{" & strT(0) & "} Error: " & ErrorDescription & "."
If Not bolNoRetry Then
If bytMaxR > 0 Then
Dim strT1 As String, i As Byte
i = Val(PrepareCol(colMaxR, s(0) & "," & s(3))) + 1
On Error Resume Next
colMaxR.Remove s(0) & "," & s(3)
On Error GoTo 0
If i <= bytMaxR Then colMaxR.Add i, s(0) & "," & s(3) Else: GoTo E
End If
If bytDelay > 0 Then
If bolDebug And Not bolLO(0) Then addLog "{" & strT(0) & "} Waiting " & bytDelay & " second(s) before another retry...", True
If bolAb Or ChkSt Then Exit Sub
SubmitIW bytDelay & "," & s(2) & "," & s(0) & "," & s(3) & "," & s(1) & "," & s(4) & ",,,,," & vbNullChar
Else
SubmitReq CInt(s(2)), CByte(s(0)), s(3), CInt(s(1)), s(4)
End If
Exit Sub
End If
E:
PrepTr s(0), s(1), s(2), s(3), True
If bolUnl Then Exit Sub
If Not tmrQ.Enabled Then ChkT Else: ChkSt
End Sub

Private Sub ChkT(Optional bolM As Boolean)
Select Case True
Case bolAb, cmdStart.Caption = "&Start", ChkSt, bytPlgUse > 0, rh.RequestCount > 0, intTmrCount > 0: Exit Sub
End Select
If Not bolM Then Enb Else: Enb 1
End Sub

Private Function ChkSt(Optional bolT As Boolean) As Boolean
If Not bolAb Then
If intAfter = 0 Then Exit Function
If Now < datCompl Then Exit Function
If Not bolT Then Enb 3
End If
ChkSt = True
End Function

Private Function ProceedString(ByVal strInp As String, strS As String, bytI As Byte, j As Byte, i As Integer, O As String, o2 As String, strTrim As String, Optional intI As Integer, Optional strI As String, Optional bolA As Boolean, Optional bytA As Byte) As String
If strInp = vbNullString Then Exit Function
Dim strT(2) As String, s() As String, strA As String, strS1 As String
If InStr(strInp, "[inp") > 0 Then
If InStr(Split(strInp, "[inp")(1), "]") > 0 Then
Select Case intI
Case -2: strT(1) = "URL"
Case -3: strT(1) = "Post"
Case -4: strT(1) = "Header name"
Case -5: strT(1) = "Header value"
Case -6: strT(1) = "If [A]"
Case -7: strT(1) = "If A <=> [B]"
Case -8: strT(1) = "Then/Else wait seconds"
Case -1: If Left$(strI, 1) = "%" Then strT(1) = strI & "%" Else: GoTo N3
Case Else: GoTo N3
End Select
strT(0) = PrepareCol(colInput, "-1," & bytI & strT(1))
If StrPtr(strT(0)) = 0 Then strT(0) = PrepareCol(colInput, j & "," & bytI & strT(1))
If StrPtr(strT(0)) <> 0 Then
s() = Split(Mid$(strT(0), 2), vbLf)
strInp = Replace(s(0), "[inp]", s(1))
strT(0) = vbNullString
End If
If strInp = vbNullString Then Exit Function
End If
End If
N3:
Dim intC(1) As Long, bolT0 As Boolean, bolT1 As Boolean, intT As Integer, intT1 As Integer, a As Byte, strT2(1) As String
bolT0 = ChkStr(strInp)
If bolAb Then Exit Function
intC(0) = 1
R:
Do
FindStr strInp, intC, bolT0, bolT1
If intC(1) = -1 Or bolAb Then Exit Function
If intC(1) = 0 Or intC(0) > intC(1) Then Exit Do
strT(0) = Mid$(strInp, intC(0), intC(1) - intC(0))
If bolT1 Then
bolT1 = False
If strT(0) <> vbNullString Then If i = 0 Then strT(2) = ProceedString(strT(0), strS, bytI, j, i, O, o2, strTrim, intT) Else: strT(2) = ProceedString(strT(0), strS, bytI, j, i, O, o2, strTrim)
intC(0) = intC(0) - 1
If InStr(intC(0) - 1, strInp, "%{" & strT(0) & "}%") > 0 Then
strInp = Left$(strInp, intC(0) - 2) & Replace(strInp, "%{" & strT(0) & "}%", "%" & strT(2) & "%", intC(0) - 1, 1)
strT(0) = strT(2)
intC(1) = InStr(intC(0), strInp, "%")
Else
C:
If bolT0 Then
intC(0) = intC(1)
GoTo R
Else: Exit Do
End If
End If
Else: strT(2) = strT(0)
End If
intC(0) = intC(0) - 1
strT(1) = PrepareCol(colStr, strT(0) & "," & bytI & "," & j & "," & O & " " & i)
If StrPtr(strT(1)) <> 0 Then strT(1) = Mid$(strT(1), InStr(strT(1), "\") + 1)
If StrPtr(strT(1)) = 0 Or Replace(strI, "%", vbNullString, , 1) = strT(0) Then
If Replace(strI, "%", vbNullString, , 1) <> strT(0) Or Not bolA And i = 0 Then
If Left$(strI, 1) <> "%" Then
If bytA > 0 Then a = bytA - 1 Else: a = 0
Do While strStrings(bytI, a) <> vbNullString
If Left$(strStrings(bytI, a), InStr(strStrings(bytI, a), vbLf) - 1) = strT(0) Then GoTo C1
a = a + 1
If UBound(strStrings, 2) < a Then Exit Do
Loop
End If
Dim strI1(2) As String, strI2(1) As String
strT(1) = vbNullString
If strI2(0) <> vbNullString Then
strI2(0) = vbNullString
strI2(1) = vbNullString
End If
strT2(0) = j & "," & Replace(Replace(O, "|", " "), "/", vbNullString, , 1)
For intT1 = colPubStr.count To 1 Step -1
strI1(0) = colPubStr.Item(intT1)
If Left$(strI1(0), InStr(strI1(0), "%") - 1) = strT(0) Then
strI1(0) = Mid$(strI1(0), Len(strT(0)) + 2)
strI1(0) = Left$(strI1(0), InStr(strI1(0), "{") - 1)
strI1(1) = Mid$(colPubStr.Item(intT1), Len(strT(0)) + Len(strI1(0)) + 3)
strI1(1) = Left$(strI1(1), InStr(InStr(strI1(1), ",") + 1, strI1(1), "\") - 1)
strI1(2) = Replace(Replace(strI1(1), "|", " "), "/", vbNullString, , 1)
Select Case True
Case InStr(strT2(0) & " ", strI1(2) & " ") = 1, InStr(" " & StrReverse(strT2(0)), " " & StrReverse(Replace(strI1(2), "...", vbNullString, , 1))) = 1, InStr(strT2(0), "...") > 0, InStr(strI1(2), "...") > 0 'If InStr(strT2(0) & " ", strI1(2) & " ") = 1 Then intT2(0) = Len(strT2(0)) - Len(strI1(2)) Else: If InStr(" " & StrReverse(strT2(0)), " " & StrReverse(Replace(strI1(2), "...", vbNullString, , 1))) = 1 Then intT2(0) = Len(strT2(0)) - Len(Replace(strI1(2), "...", vbNullString, , 1)) Else: If InStr(strT2(0), "...") > 0 Or InStr(strI1(2), "...") > 0 Then intT2(0) = -Len(strI1(1)) Else: intT2(0) = -1
strI2(0) = strI1(0)
strI2(1) = strI1(1)
strT(1) = Mid$(colPubStr.Item(intT1), Len(strT(0)) + Len(strI1(0)) + Len(strI1(1)) + 4)
Exit For
End Select
End If
N2:
If bolAb Then Exit Function
Next
If strT(1) <> vbNullString Then
If strI2(0) = bytI Then strI = strT(0)
intT1 = Left(strT(1), InStr(strT(1), "\") - 1)
strT(1) = Mid$(strT(1), InStr(strT(1), "\") + 1)
If intT1 > 0 Then
strA = bytI & "," & O
strS1 = strS
bytI = CByte(strI2(0))
O = Mid$(strI2(1), InStr(strI2(1), ",") + 1)
If Left$(strT(1), 1) <> "," Then O = TrimO(O & " " & Left$(strT(1), InStr(strT(1), ",") - 1))
strS = vbNullString
GetSrc strS, CStr(j), O, bytI
End If
strT(1) = Mid$(strT(1), InStr(strT(1), ",") + 1)
If strI <> vbNullString Or intT1 > 0 Then strI = "{" & strI
End If
End If
If bytA > 0 Then a = bytA - 1 Else: a = 0
C1:
If StrPtr(strT(1)) = 0 Then
strT(1) = PrepareCol(colInput, "-1," & bytI & "%" & strT(0) & "%")
If StrPtr(strT(1)) = 0 Then strT(1) = PrepareCol(colInput, j & "," & bytI & "%" & strT(0) & "%")
If StrPtr(strT(1)) <> 0 Then
s() = Split(Mid$(strT(1), 2), vbLf)
If Left$(strT(1), 1) = "1" Or Left$(strT(1), 1) = "3" Then
If i = 0 Then If intT = 0 Or intT > UBound(s()) Then intT = UBound(s())
strT(1) = Replace(s(0), "[inp]", s(i + 1))
Else: strT(1) = Replace(s(0), "[inp]", s(1))
End If
strI = "{"
End If
End If
If Replace(Replace(strI, "%", vbNullString, , 1), "{", vbNullString, , 1) = strT(0) And strA = vbNullString Then GoTo N
Do While strStrings(bytI, a) <> vbNullString
If Split(strStrings(bytI, a), vbLf)(0) = strT(0) Then
If Left$(strI, 1) <> "{" Or strT(1) <> vbNullString And intT1 > 0 And i > 0 Then strT(1) = Split(strStrings(bytI, a), vbLf)(1)
If strA <> vbNullString And InStr(Split(strT(1) & "[inp", "[inp")(1), "]") > 0 Then strT(1) = "%" & strT(0) & "%"
s() = Split(Split(strStrings(bytI, a), vbLf)(2) & ",,,", ",")
If s(2) = "1" Then
If i = 0 Then
If strT(1) = vbNullString Or intT1 = 0 Then strT2(0) = ProceedString(strT(1), strS, bytI, j, 0, O, o2, strTrim, intT, "%" & strT(0), True)
Else: strT2(0) = ProceedString(strT(1), strS, bytI, j, i, O, o2, strTrim, , "%" & strT(0))
End If
Else: strT2(0) = ProceedString(strT(1), strS, bytI, j, i, O, o2, strTrim, -1, "%" & strT(0))
End If
If intI <> -1 And strA = vbNullString Then If intT1 > 0 Then If intT = 0 Or intT > intT1 Then intT = intT1
If strT(1) <> vbNullString And intT1 > 0 And i = 0 Or bolAb Then
Ex:
If strA <> vbNullString Then
s() = Split(strA, ",")
strA = vbNullString
bytI = s(0)
O = s(1)
strS = strS1
End If
Else
If strT2(0) = vbNullString And s(0) = "1" Then
If Not bolLO(0) Then addLog "{T: " & j + 1 & ", S: " & i & ", I: " & bytI + 1 & ", O:" & o2 & "} Error: Crucial string %" & strT(0) & "% is blank."
bolT1 = True
GoTo Ex
End If
If strT(1) <> strT2(0) Then
If bolDebug And Not bolLO(0) Then If intT = -1 Or intT = 0 Then addLog "{T: " & j + 1 & ", S: " & i & ", I: " & bytI + 1 & ", O:" & o2 & "} %" & strT(0) & "%: " & strT2(0), True
strT(1) = strT2(0)
End If
If bolAb Then Exit Function
If strT(1) <> vbNullString And intT1 > 0 Then GoTo Ex
If s(1) = "1" Then
If s(2) <> "1" Or i > 0 Then strT2(0) = strTrim Else: strT2(0) = O
If strT2(0) <> O Then strT2(1) = "," & strT(1) Else: strT2(1) = i & "," & strT(1)
On Error Resume Next
colPubStr.Remove strT(0) & "%" & j & "," & strT2(0)
On Error GoTo 0
colPubStr.Add strT(0) & "%" & bytI & "{" & j & "," & strT2(0) & "\" & intT & "\" & strT2(1), strT(0) & "%" & j & "," & strT2(0)
End If
colStr.Add bytI & "," & j & "," & O & " " & i & "\" & strT(1), strT(0) & "," & bytI & "," & j & "," & O & " " & i
If Not bolLO(1) Then
If s(3) = "1" Then
If intOutMax > 0 Then If (Len(txtOutput.Text) - Len(Replace(txtOutput.Text, vbNewLine, vbNullString))) / 2 = intOutMax Then txtOutput.Text = Mid$(txtOutput.Text, InStr(txtOutput.Text, vbNewLine) + 2)
txtOutput.Text = txtOutput.Text & Replace(Replace(Replace(Replace(Replace(strTemplate0, "{T}", j + 1), "{S}", i), "{O}", o2), "{I}", bytI + 1), "{N}", strT(0)) & strT(1) & strTemplate1
txtOutput.ScrollToBottom
End If
End If
End If
If bolT1 Then Exit Function
Exit Do
End If
a = a + 1
If bolAb Then Exit Function
If UBound(strStrings, 2) < a Then Exit Do
Loop
If bolT0 Then
If StrPtr(strT(1)) <> 0 Then GoTo N Else: intC(0) = intC(1) + 1
Else: GoTo N
End If
Else
N:
If Left$(strI, 1) = "{" Then strI = Mid$(strI, 2)
If Not bolT0 Then
strT(1) = Replace(strT(1), "'", "''")
If strT(1) Like "*[!0-9]*" Then strT(1) = "'" & strT(1) & "'"
End If
strInp = Left$(strInp, intC(0) - 1) & Replace(strInp, "%" & strT(0) & "%", strT(1), intC(0), 1)
intC(0) = intC(0) + Len(strT(1))
End If
If bolAb Then Exit Function
Loop
If bytA > 0 Then
ProceedString = ""
Exit Function
End If
If InStr(strInp, "[oind]") > 0 And O <> vbNullString Then
strT(0) = Mid$(O, InStrRev(O, " ") + 1)
strInp = Replace(strInp, "[oind]", Mid$(strT(0), InStr(strT(0), "x") + 1))
End If
strInp = Replace(strInp, "[cind]", i)
strInp = Replace(strInp, "[thr]", j + 1)
If bolT0 Then
ProceedString = ReplaceString(strInp, strS)
ElseIf Not IsMissing(intI) Then
If intI > -1 Then
If i = 0 Then
If bolA Then ProceedString = ProcessString(ReplaceString(strInp), strS, , , intT) Else: ProceedString = ProcessString(ReplaceString(strInp), strS)
Else: ProceedString = ProcessString(ReplaceString(strInp), strS, , i)
End If
Else: ProceedString = ProcessString(ReplaceString(strInp), strS)
End If
Else: ProceedString = ProcessString(ReplaceString(strInp), strS)
End If
If StrPtr(ProceedString) = 0 Then ProceedString = ""
If intT > 0 Then If intT < intI Or intI < 1 Then intI = intT
If intI < 1 Then Exit Function
Dim strM As String: strM = PrepareCol(colMax, bytI & "," & j & "," & O)
If strM <> vbNullString Then If Val(strM) <= intI - 1 Then Exit Function
On Error Resume Next
colMax.Remove bytI & "," & j & "," & O
colMax.Add intI - 1, bytI & "," & j & "," & O
End Function

Private Sub CheckIf(s0 As String, s1 As String, s2 As String, s3 As String, strS As String, strT1 As String, strTrim As String, Optional bolEn As Boolean)
Dim strT2 As String, bolA As Boolean, strT(1) As String
bolA = PrepareCol(colMax, s2 & "," & s0 & "," & s3) = vbNullString
If Left$(s3, 1) = "/" Or InStr(strT1, "-") > 0 Or InStr(strT1, "+") > 0 Then strT2 = strT1
If strIf(s2, 0) <> vbNullString Then
Dim i As Byte, s() As String, ns(1) As Byte, con(1) As Variant
For i = 0 To UBound(strIf, 2)
If strIf(s2, i) <> vbNullString Then
s() = Split(strIf(s2, i), vbLf)
If i > 0 Then If ns(0) = 0 And s(3) = "0" Then Exit For Else: If ns(0) = 1 And s(3) = "1" Then GoTo N
If bolDebug Then strT(0) = s(0): strT(1) = s(2)
con(0) = ProceedString(s(0), strS, CByte(s2), CByte(s0), CInt(s1), s3, strT2, strTrim, -6)
con(1) = ProceedString(s(2), strS, CByte(s2), CByte(s0), CInt(s1), s3, strT2, strTrim, -7)
If bolAb Then Exit Sub
If bolDebug And Not bolLO(0) Then
If con(0) <> s(0) Then addLog "{T: " & s0 + 1 & ", S: " & s1 & ", I: " & s2 + 1 & ", O:" & strT2 & "} If A (" & i & "): " & con(0), True
If con(1) <> s(2) Then addLog "{T: " & s0 + 1 & ", S: " & s1 & ", I: " & s2 + 1 & ", O:" & strT2 & "} If A <=> [B] (" & i & "): " & con(1), True
End If
If IsNumeric(con(0)) Then con(0) = Val(con(0))
If IsNumeric(con(1)) Then con(1) = Val(con(1))
ns(0) = 1
Select Case s(1)
Case 0: If con(0) = con(1) Then ns(0) = 0
Case 1: If con(0) <> con(1) Then ns(0) = 0
Case 2: If con(0) > con(1) Then ns(0) = 0
Case 3: If con(0) < con(1) Then ns(0) = 0
Case 4: If con(0) >= con(1) Then ns(0) = 0
Case 5: If con(0) <= con(1) Then ns(0) = 0
End Select
ElseIf i > bytSh(s2) Then Exit For
End If
N:
Next
End If
Dim sp() As String
i = 0
Do While strStrings(s2, i) <> vbNullString
s() = Split(strStrings(s2, i), vbLf)
sp() = Split(s(2) & ",,,", ",")
If sp(1) = "1" Or sp(3) = "1" Then
If sp(2) = "1" And s1 = "0" Then strT(0) = s3 Else: strT(0) = s3 & " " & s1
If PrepareCol(colStr, s(0) & "," & s2 & "," & s0 & "," & s3 & " " & s1) = vbNullString And PrepareCol(colPubStr, s(0) & "%" & s0 & "," & strT(0)) = vbNullString Then
If StrPtr(ProceedString("%" & s(0) & "%", strS, CByte(s2), CByte(s0), CInt(s1), s3, strT2, strTrim, , , , i + 1)) = 0 Then GoTo E
End If
End If
i = i + 1
If bolAb Then Exit Sub
If UBound(strStrings, 2) < i Then Exit Do
Loop
If ns(0) = 0 Then
If intGoto(0, s2) > 1 Then
ns(1) = intGoto(0, s2) - 2
ElseIf intGoto(0, s2) = 0 Then ns(1) = s2 + 1
Else
E:
PrepTr s0, s1, s2, s3
If bolUnl Then Exit Sub
If Not tmrQ.Enabled Then ChkT Else: ChkSt
Exit Sub
End If
ElseIf intGoto(1, s2) > 0 Then
If intGoto(1, s2) = 1 Then ns(1) = s2 + 1 Else: ns(1) = intGoto(1, s2) - 2
Else: GoTo E
End If
Dim N As Integer, strT3 As String
N = ns(1) - s2
If strWait(ns(0), s2) <> vbNullString Then
strT(0) = ProceedString(strWait(ns(0), s2), strS, CByte(s2), CByte(s0), CInt(s1), s3, strT2, strTrim, -8)
If bolAb Or ChkSt Then Exit Sub
If IsNumeric(strT(0)) Then
If bolDebug Then
If ns(0) = 0 Then strT(1) = "Then" Else: strT(1) = "Else"
If Not bolLO(0) Then If strT(0) <> strWait(ns(0), s2) Then addLog "{T: " & s0 + 1 & ", S: " & s1 & ", I: " & s2 + 1 & ", O:" & strT2 & "} " & strT(1) & " wait: " & strT(0), True
End If
GoSub TrO
SubmitIW Replace(strT(0), ",", ".") & "," & s0 & "," & s1 & "," & s2 & "," & s3 & "," & ns(1) & "," & strT1 & "," & strT3 & "," & CInt(bolA) & "," & CInt(bolEn) & "," & strS
Exit Sub
End If
End If
GoSub TrO
Finish s0, s1, s2, s3, ns(1), strS, strT1, strT3, bolA, bolEn
Exit Sub
TrO:
PrepTr s0, s1, s2, s3, bolEn, strS, N, strT3
If N <> 0 Then
Dim strL As String, l As Integer
l = InStrRev(strT1, " ")
strL = "x" & Mid$(strT1, l + 1)
If Mid$(strL, InStrRev(strL, "x") + 1) = s1 Then
strL = Mid$(strL, 2)
If InStr(strL, "x") > 0 Then
strT1 = Left$(strT1, l) & Left$(strL, InStr(strL, "x") - 1) + 1 & "x" & s1
Else: strT1 = Left$(strT1, l) & "2x" & s1
End If
Else: strT1 = strT1 & " " & s1
End If
End If
If N > 1 Then strT1 = strT1 & " +" & N Else: If N < 0 Then strT1 = strT1 & " " & N
If bytTOrigin0 > 0 Then
Dim lngC As Long: lngC = Len(strT1) - Len(Replace(strT1, " ", vbNullString)) - bytTOrigin0
If lngC > 0 Then
strT1 = Replace(strT1, " ", vbNullString, , lngC)
strT1 = " ..." & Mid$(strT1, InStr(strT1, " "))
End If
End If
Return
End Sub

Private Sub PrepTr(s0 As String, s1 As String, s2 As String, s3 As String, Optional bolEn As Boolean, Optional strS As String, Optional N As Integer, Optional strT3 As String)
Dim a As Integer, s31 As String, strT1 As String
s31 = s3 & " " & s1
Static strT As String
strT = strT & s2 & "," & s0 & "," & s31 & "\" & vbLf
If Len(strT) - Len(Replace(strT, vbLf, vbNullString)) = 1 Then
Do While strT <> vbNullString
strT1 = Left$(strT, InStr(strT, vbLf) - 1)
For a = colStr.count To 1 Step -1
If InStr(colStr.Item(a), strT1) = 1 Then colStr.Remove a
Next
strT = Mid$(strT, Len(strT1) + 2)
Loop
End If
If bytOrigin > 0 Then
On Error Resume Next
colCurrO.Remove s3
On Error GoTo 0
End If
strT3 = TrimO(s31, s0, N)
If bolEn Or strS = vbNullString Then Exit Sub
If bytOrigin = 0 Then
On Error Resume Next
colSrc.Remove s0 & "," & s2 & "," & strT3
On Error GoTo 0
colSrc.Add strS, s0 & "," & s2 & "," & strT3
Else
On Error Resume Next
colSrc.Remove s0 & "," & strT3
On Error GoTo 0
colSrc.Add s0 & "," & strT3 & "\" & strS, s0 & "," & strT3
End If
End Sub

Private Sub SubmitIW(strT As String, Optional typ As Byte)
Dim obj As Object
If typ = 1 Then Set obj = tmrI Else: Set obj = tmrW
If obj(0).Enabled Then
Dim i As Integer
If intLTmr(typ) > 0 Then i = intLTmr(typ) Else: i = 1
On Error GoTo E
C:
If bolAb Then Exit Sub
Load obj(i)
If intLTmr(typ) = 0 Then intLTmr(typ) = i + 1
End If
If bytOrigin > 0 Then
Dim strT1 As String
strT1 = Mid$(strT, InStr(InStr(InStr(InStr(strT, ",") + 1, strT, ",") + 1, strT, ",") + 1, strT, ",") + 1)
On Error Resume Next
colCurrO.Add "", Left$(strT1, InStr(strT1, ",") - 1)
On Error GoTo 0
End If
obj(i).Tag = strT
intTmrCount = intTmrCount + 1
obj(i).Enabled = True
Set obj = Nothing
Exit Sub
E:
i = i + 1
Resume C
End Sub

Private Sub Finish(s0 As String, s1 As String, s2 As String, s3 As String, ns As Byte, strS As String, strT1 As String, strT As String, bolA As Boolean, bolEn As Boolean)
Dim bolM As Boolean
If ns > bytIC Then
bolM = True
If bolEn Then
bolEn = bolM
Else
If bolUnl Then Exit Sub
If Not tmrQ.Enabled Then ChkT bolM Else: ChkSt
End If
Else
Dim strM As String
If bolA Then
strM = PrepareCol(colMax, s2 & "," & s0 & "," & s3)
If strM <> vbNullString Then
intLTmr(0) = 0
If bytOrigin <> 0 Then colMax.Remove s2 & "," & s0 & "," & s3
End If
End If
Dim a As Integer
For a = 0 To Val(strM)
If bolAb Or ChkSt Then Exit Sub
If strURLData(ns) = vbNullString Then
SubmitIW s0 & "," & a & "," & ns & "," & strT & "," & strT1 & "," & TrimO(strT & " " & a) & "," & strS, 1
Else: SubmitReq ns, CByte(s0), strT, CInt(a), strT1
End If
Next
End If
End Sub

Private Function TrimO(strT As String, Optional j As String, Optional N As Integer) As String
Dim b As Byte, lngC As Long
If bytOrigin > 0 Then
lngC = Len(strT) - Len(Replace(strT, " ", vbNullString)) - bytOrigin
If lngC = bytOrigin Then
Static bytT(1) As Byte, colR(1) As Collection, strC(1) As String
If j <> vbNullString Then
bytT(0) = bytT(0) + 1
If bytT(0) = 1 Then Set colR(0) = New Collection
Dim strT1(7) As String
End If
strT1(0) = Replace(strT, " ", vbNullString, , bytOrigin)
lngC = InStr(strT1(0), " ")
If Left$(strT, 1) = "/" Then TrimO = "/"
TrimO = TrimO & "..." & Mid$(strT1(0), lngC)
If j <> vbNullString Then
Dim a As Integer, i As Integer, colT As Collection, varT As Variant
strT1(0) = Left$(strT, lngC + bytOrigin)
Set colT = New Collection
Dim obj As Collection
Set obj = colPubStr
R:
For i = obj.count To 1 Step -1
strT1(1) = obj.Item(i)
PullK b = 0, strT1(1), strT1(7), a
If StrPtr(PrepareCol(colR(b), strT1(7))) = 0 Then
If a > 1 Then strT1(2) = Mid$(strT1(7), InStr(strT1(7), "%") + 1) Else: strT1(2) = strT1(7)
If Left$(strT1(2), InStr(strT1(2), ",") - 1) = j Then
strT1(2) = Left$(strT1(1), InStr(a, strT1(1), ","))
strT1(3) = Mid$(strT1(1), Len(strT1(2)) + 1)
If Left$(strT1(3), 1) = "/" Then strT1(6) = "/" Else: strT1(6) = vbNullString
strT1(4) = Left$(strT1(3), InStr(strT1(3), "\") - 1) & " "
If StrPtr(PrepareCol(colCurrO, Left$(strT1(4), Len(strT1(4)) - 1))) = 0 Then
If InStr(tmrQ.Tag, "," & j & "," & Left$(strT1(4), Len(strT1(4)) - 1) & ",") = 0 Then
If strT1(4) <> strT1(6) & "... " Then
If Len(strT1(4)) >= Len(strT1(6)) + Len(strT1(0)) And Left$(strT1(4), Len(strT1(6)) + 3) = Left$(strT1(0), Len(strT1(6)) + 3) Then
lngC = InStr(Replace(strT1(4), "|", " "), Replace(strT1(0), "|", " "))
If lngC = 1 Or lngC = 2 Then
If lngC = 2 Then strT1(5) = strT1(6) & strT1(0) Else: strT1(5) = strT1(0)
strT1(5) = Replace(strT1(4), strT1(5), strT1(6) & "... ", , 1)
strT1(5) = Left$(strT1(5), Len(strT1(5)) - 1)
Else: GoTo N
End If
Else
lngC = InStr(Replace(strT1(0), "|", " "), Replace(strT1(4), "|", " "))
If lngC = 1 Or lngC = 2 Then strT1(5) = strT1(6) & "..." Else: GoTo N
End If
strT1(1) = Left$(strT1(1), Len(strT1(2))) & strT1(5)
If strT1(5) = strT1(6) & "..." Then
For Each varT In colT
If InStr(varT, strT1(2)) = 1 Then GoTo N1
Next
strC(b) = strC(b) & strT1(2) & vbLf
ElseIf PrepareCol(colT, strT1(1)) <> vbNullString Then GoTo N
End If
colT.Add strT1(1) & Mid$(strT1(3), Len(strT1(4))), strT1(1)
N1:
On Error Resume Next
colR(b).Add strT1(7), strT1(7)
On Error GoTo 0
ElseIf InStr(vbLf & strC(b), vbLf & strT1(2) & vbLf) > 0 Then GoTo N1
Else: strC(b) = strC(b) & strT1(2) & vbLf
End If
End If
End If
N:
End If
End If
Next
On Error Resume Next
For i = colT.count To 1 Step -1
strT1(1) = colT.Item(i)
PullK b = 0, strT1(1), strT1(2)
obj.Remove strT1(2)
obj.Add strT1(1), strT1(2), obj.count
colR(b).Remove strT1(2)
Next
On Error GoTo 0
Set colT = Nothing
bytT(b) = bytT(b) - 1
i = 1
Do While bytT(b) = 0 And colR(b).count >= i
obj.Remove colR(b).Item(i)
i = i + 1
Loop
If bytT(b) = 0 Then Set colR(b) = Nothing: strC(b) = vbNullString
If b = 0 Then
Set obj = colSrc
Set colT = New Collection
If bytT(1) = 0 Then Set colR(1) = New Collection
b = 1
a = 1
bytT(1) = bytT(1) + 1
GoTo R
End If
End If
GoTo E
End If
End If
If j <> vbNullString Then If Left$(strT, 1) <> "/" Then If Mid$(strT, InStrRev(strT, " ") + 1, 1) <> "0" Then TrimO = "/"
TrimO = TrimO & strT
E:
If N < 0 Then
If bolColl Then
lngC = Len(TrimO)
For b = 1 To Abs(N)
lngC = InStrRev(TrimO, " ", lngC - 1)
If lngC = 0 Then
TrimO = "..."
GoTo C
End If
Next
TrimO = Left$(TrimO, lngC - 1)
lngC = InStrRev(Mid$(TrimO, InStrRev(TrimO, " ") + 1), "|")
If lngC > 0 Then TrimO = Left$(TrimO, InStrRev(TrimO, " ") + lngC - 1)
End If
C: TrimO = TrimO & "|" & N
ElseIf N > 1 Then TrimO = TrimO & "|+" & N
End If
End Function

Private Sub PullK(bolT As Boolean, strI As String, strO As String, Optional a As Integer)
Dim strT(1) As String
If bolT Then
strT(0) = Left$(strI, InStr(strI, "%") - 1)
a = InStr(Len(strT(0)), strI, "{") + 1
strT(1) = Mid$(strI, a)
strO = strT(0) & "%" & Left$(strT(1), InStr(strT(1), "\") - 1)
Else: strO = Left$(strI, InStr(strI, "\") - 1)
End If
End Sub

Private Sub tmrQ_Timer()
Dim s(4) As String, i As Byte, ln(1) As Long, strT As String
If Not bolAb And tmrQ.Tag <> vbNullString Then
If 255 - bytActive < 1 Then Exit Sub
strT = Mid$(tmrQ.Tag, 2, InStr(tmrQ.Tag, vbLf) - 2)
ln(0) = 1
For i = 0 To 4
ln(1) = InStr(ln(0), strT, ",")
If ln(1) = 0 Then ln(1) = Len(strT) + 1
s(i) = Mid$(strT, ln(0), ln(1) - ln(0))
ln(0) = ln(1) + 1
Next
Dim intT As Integer
If s(3) = "0" And Left$(s(3), 1) <> "/" Then intT = 255 Else: intT = intSubT
If intT - bytActive > 0 Then SubmitReq CByte(s(0)), CByte(s(1)), s(2), CInt(s(3)), s(4) Else: Exit Sub
If tmrQ.Tag <> vbNullString And Not bolAb Then Exit Sub
End If
tmrQ.Enabled = False
tmrQ.Tag = vbNullString
If bolUnl Then Exit Sub
If ChkSt Then Exit Sub
ChkT
End Sub

Private Sub tmrW_Timer(Index As Integer)
If Not bolAb Then
If Left$(tmrW(Index).Tag, InStr(tmrW(Index).Tag, ",") - 1) > 1 Then
Debug.Print Left$(tmrW(Index).Tag, InStr(tmrW(Index).Tag, ",") - 1) - 1
If ChkSt(True) Then
tmrW(Index).Enabled = False
intTmrCount = intTmrCount - 1
If Index > 0 Then Unload tmrW(Index)
Enb 3
Exit Sub
End If
tmrW(Index).Tag = Left$(tmrW(Index).Tag, InStr(tmrW(Index).Tag, ",") - 1) - 1 & Mid$(tmrW(Index).Tag, InStr(tmrW(Index).Tag, ","))
Exit Sub
End If
tmrW(Index).Enabled = False
Dim s(9) As String, i As Byte, ln(1) As Long
ln(0) = InStr(tmrW(Index).Tag, ",") + 1
For i = 0 To 8
ln(1) = InStr(ln(0), tmrW(Index).Tag, ",")
s(i) = Mid$(tmrW(Index).Tag, ln(0), ln(1) - ln(0))
ln(0) = ln(1) + 1
Next
s(9) = Mid$(tmrW(Index).Tag, ln(0))
If s(9) <> vbNullChar Then
intTmrCount = intTmrCount - 1
Finish s(0), s(1), s(2), s(3), CByte(s(4)), s(9), s(5), s(6), CBool(s(7)), CBool(s(8))
If bolUnl Then Exit Sub
Else
intTmrCount = intTmrCount - 1
SubmitReq CByte(s(0)), CByte(s(1)), s(2), CInt(s(3)), s(4)
End If
Else
intTmrCount = intTmrCount - 1
If Not bolUnl Then tmrW(Index).Enabled = False
End If
If Not bolUnl Then If Index > 0 Then Unload tmrW(Index): intLTmr(1) = Index
If Not bolAb Then If s(9) <> vbNullChar Then If Not tmrQ.Enabled Then ChkT CBool(s(7)) Else: ChkSt
End Sub

Private Sub tmrU_Timer(Index As Integer)
tmrU(Index).Enabled = False
intTmrCount = intTmrCount - 1
If Not bolAb Then SubmitReq 0, CByte(Index)
If Index > 0 Then Unload tmrU(Index)
End Sub

Private Sub tmrI_Timer(Index As Integer)
tmrI(Index).Enabled = False
If Not bolAb Then
Dim s(6) As String
If tmrI(Index).Tag <> vbNullString Then
Dim i As Byte, ln(1) As Long
ln(0) = 1
For i = 0 To 5
ln(1) = InStr(ln(0), tmrI(Index).Tag, ",")
s(i) = Mid$(tmrI(Index).Tag, ln(0), ln(1) - ln(0))
ln(0) = ln(1) + 1
Next
s(6) = Mid$(tmrI(Index).Tag, ln(0))
Else
s(0) = Index
s(1) = "0"
s(2) = "0"
End If
Dim bolM As Boolean: bolM = True
intTmrCount = intTmrCount - 1
CheckIf s(0), s(1), s(2), s(3), s(6), s(4), s(5), bolM
Else: intTmrCount = intTmrCount - 1
End If
If bolUnl Then Exit Sub
If Index > 0 Then
Unload tmrI(Index)
intLTmr(0) = Index
End If
If Not tmrQ.Enabled Then ChkT bolM Else: ChkSt
End Sub

Private Function PrepareCol(col As Collection, k As String) As String
On Error Resume Next
PrepareCol = col.Item(k)
End Function

Private Function ChkStr(strInp As String) As Boolean
Dim intC(1) As Long
If Left$(strInp, 1) = "[" Then
ChkStr = True
Exit Function
ElseIf Left$(strInp, 1) = "'" And InStr(2, strInp, "'") > 0 Then
intC(0) = FindC1(strInp)
If intC(0) <> Len(strInp) Then If Mid$(strInp, intC(0) + 1, 1) <> "+" Then ChkStr = True
Exit Function
Else: If Left$(strInp, 1) = "%" Then If Mid$(strInp, 2, 1) <> "{" Then If InStr(2, strInp, "%") = Len(strInp) Then Exit Function
End If
If InStr("%<", Left$(strInp, 1)) = 0 Then
If InStr(strInp, "(") > 0 Then
If InStr(Comms & strPlC, "," & Left$(strInp, InStr(strInp, "(") - 1) & ",") = 0 Then ChkStr = True
Else: ChkStr = True
End If
ElseIf Left$(strInp, 1) = "<" Then
ChkStr = Mid$(strInp & "+", InStr(2, strInp, ">") + 1, 1) <> "+"
Exit Function
Else
Dim bytT As Byte
intC(0) = 1
FindStr strInp, intC
If intC(1) <= 0 Or intC(0) > intC(1) Then Exit Function
If Mid$(strInp, intC(1), 1) = "}" Then bytT = 2 Else: bytT = 1
ChkStr = Mid$(strInp, intC(1) + bytT, 1) <> "+"
End If
End Function

Private Sub FindStr(strInp As String, intC() As Long, Optional bolT0 As Boolean, Optional bolT1 As Boolean)
If Not bolT0 Then
Do
intC(0) = FindSep(strInp, intC(0), "%") + 1
If intC(0) = 1 Then intC(1) = 0: Exit Sub
intC(1) = FindSep(strInp, intC(0), ">") + 1
Loop Until intC(1) = 1 Or intC(1) < intC(0)
Else: intC(0) = InStr(intC(0), strInp, "%") + 1
End If
If intC(0) = 1 Then intC(1) = 0: Exit Sub
Dim a As Byte
If Mid$(strInp, intC(0), 1) = "{" Then
bolT1 = True
intC(0) = intC(0) + 1
Dim intP(1) As Long
intP(0) = intC(0)
intP(1) = intC(0)
Do
intP(1) = FindSep(strInp, intP(1), "}") + 1
If intP(1) = 1 Then
If Not bolT0 Then intC(1) = -1
Exit Sub
End If
intC(1) = intP(1) - 1
intP(0) = FindSep(strInp, intP(0), "{") + 1
If bolAb Then Exit Sub
Loop Until intP(0) = 1 Or intP(0) > intP(1)
Else: intC(1) = InStr(intC(0), strInp, "%")
End If
End Sub

Private Sub cmdSave_Click(Index As Integer)
If Index = 0 Then
If lstLog.list(0) = vbNullString Then Exit Sub
ElseIf txtOutput.Text = vbNullString Then Exit Sub
End If
Dim strT(2) As String
strT(1) = "Text file (*.txt)|*.txt"
strT(0) = Replace(Replace(Now, "/", "."), ":", "-")
If Index = 0 Then
If Left$(strPath(0), InStr(strPath(0) & ".", ".") - 1) = "{NOW}" Then strPath(0) = strT(0) & ".log"
strT(2) = "log"
strT(1) = "Log file (*.log)|*.log|" & strT(1)
Else
If Left$(strPath(1), InStr(strPath(1) & ".", ".") - 1) = "{NOW}" Then strPath(1) = strT(0) & ".txt"
strT(0) = "output"
strT(1) = strT(1) & "|Custom type (*.*)|*.*"
strT(2) = strT(0)
End If
Dim strFile As String
If strPath(Index) = vbNullString Then
strFile = CommDlg(True, "Select where to save " & strT(2), strT(1), , strT(0))
If strFile = vbNullString Then Exit Sub
Else: strFile = strPath(Index)
End If
On Error GoTo E
lblStatus.Caption = "Saving " & strT(2) & "..."
lblStatus.Refresh
If Index = 0 Then
Open strFile For Output Access Write As #1
strT(1) = vbNullString
Dim i As Integer
For i = 0 To lstLog.ListCount - 1
strT(1) = strT(1) & lstLog.list(i) & vbCrLf
Next
strT(1) = Left$(strT(1), Len(strT(1)) - 2)
Print #1, strT(1);
Close #1
Else: PutContents strFile, txtOutput.Text, IIf(IsUnicode(txtOutput.Text), CP_UTF16_LE, CP_ACP)
End If
If bolDebug And Not bolLO(0) Then addLog UCase$(Left$(strT(2), 1)) & Mid$(strT(2), 2) & " saved (" & get_relative_path_to(strFile) & ").", True
lblStatus.Caption = "Idle..."
Exit Sub
E:
Close #1
If bolDebug And Not bolLO(0) Then addLog "Failed to save " & strT(2) & ".", True
lblStatus.Caption = "Error! Idle..."
If strPath(Index) = vbNullString Then MsgBox "Error in saving " & strT(2) & " file!", vbCritical
End Sub

Private Sub cmdClear_Click(Index As Integer)
If MsgBox("Sure?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
If Index = 0 Then
lstLog.Clear
SetListboxScrollbar
Else: txtOutput.Text = vbNullString
End If
cmdClear(Index).Enabled = False
cmdSave(Index).Enabled = False
End Sub

Public Function IsUnicode(s As String) As Boolean
   Dim i As Long
   Dim bLen As Long
   Dim Map() As Byte

   If LenB(s) Then
      Map = s
      bLen = UBound(Map)
      For i = 1 To bLen Step 2
         If (Map(i) > 0) Then
            IsUnicode = True
            Exit Function
         End If
      Next
   End If
End Function

Private Sub WriteBOM(ByVal the_iFileNo As Integer, ByVal the_nCodePage As Long)

    ' FF FE         UTF-16, little endian
    ' FE FF         UTF-16, big endian
    ' EF BB BF      UTF-8
    ' FF FE 00 00   UTF-32, little endian
    ' 00 00 FE FF   UTF-32, big-endian

    Select Case the_nCodePage
    Case CP_UTF8
        Put #the_iFileNo, , CByte(&HEF)
        Put #the_iFileNo, , CByte(&HBB)
        Put #the_iFileNo, , CByte(&HBF)
    Case CP_UTF16_LE
        Put #the_iFileNo, , CByte(&HFF)
        Put #the_iFileNo, , CByte(&HFE)
    Case CP_UTF16_BE
        Put #the_iFileNo, , CByte(&HFE)
        Put #the_iFileNo, , CByte(&HFF)
    Case CP_UTF32_LE
        Put #the_iFileNo, , CByte(&HFF)
        Put #the_iFileNo, , CByte(&HFE)
        Put #the_iFileNo, , CByte(&H0)
        Put #the_iFileNo, , CByte(&H0)
    Case CP_UTF32_BE
        Put #the_iFileNo, , CByte(&H0)
        Put #the_iFileNo, , CByte(&H0)
        Put #the_iFileNo, , CByte(&HFE)
        Put #the_iFileNo, , CByte(&HFF)
    End Select

End Sub

' Purpose:  Analogue of 'Open "fileName" For Output As #fileNo'
Private Sub OpenForOutput(ByRef the_sFilename As String, ByVal the_iFileNo As Integer, Optional ByVal the_nCodePage As Long = CP_ACP, Optional ByVal the_bPrefixWithBOM As Boolean = True)

    ' Ensure we overwrite the file by deleting it ...
    On Error Resume Next
    Kill the_sFilename
    On Error GoTo 0

    ' ... before creating it.
    Open the_sFilename For Binary Access Write As #the_iFileNo

    If the_bPrefixWithBOM Then
        WriteBOM the_iFileNo, the_nCodePage
    End If

End Sub

' Purpose:  Analogue of the 'Print #fileNo, value' statement. But only one value allowed.
'           Setting <the_bAppendNewLine> = False is analagous to 'Print #fileNo, value;'.
Private Sub Print_(ByVal the_iFileNo As Integer, ByRef the_sValue As String, Optional ByVal the_nCodePage As Long = CP_ACP, Optional ByVal the_bAppendNewLine As Boolean = True)

    Const kbytNull                  As Byte = 0
    Const kbytCarriageReturn        As Byte = 13
    Const kbytNewLine               As Byte = 10

    Put #the_iFileNo, , ToCPString(the_sValue, the_nCodePage)

    If the_bAppendNewLine Then
        Select Case the_nCodePage
        Case CP_UTF16_BE
            Put #the_iFileNo, , kbytNull
            Put #the_iFileNo, , kbytCarriageReturn
            Put #the_iFileNo, , kbytNull
            Put #the_iFileNo, , kbytNewLine
        Case CP_UTF16_LE
            Put #the_iFileNo, , kbytCarriageReturn
            Put #the_iFileNo, , kbytNull
            Put #the_iFileNo, , kbytNewLine
            Put #the_iFileNo, , kbytNull
        Case CP_UTF32_BE
            Put #the_iFileNo, , kbytNull
            Put #the_iFileNo, , kbytNull
            Put #the_iFileNo, , kbytNull
            Put #the_iFileNo, , kbytCarriageReturn
            Put #the_iFileNo, , kbytNull
            Put #the_iFileNo, , kbytNull
            Put #the_iFileNo, , kbytNull
            Put #the_iFileNo, , kbytNewLine
        Case CP_UTF32_LE
            Put #the_iFileNo, , kbytCarriageReturn
            Put #the_iFileNo, , kbytNull
            Put #the_iFileNo, , kbytNull
            Put #the_iFileNo, , kbytNull
            Put #the_iFileNo, , kbytNewLine
            Put #the_iFileNo, , kbytNull
            Put #the_iFileNo, , kbytNull
            Put #the_iFileNo, , kbytNull
        Case Else
            Put #the_iFileNo, , kbytCarriageReturn
            Put #the_iFileNo, , kbytNewLine
        End Select
    End If

End Sub

Private Sub PutContents(ByRef the_sFilename As String, ByRef the_sFileContents As String, ByVal the_nCodePage As Long, Optional the_bPrefixWithBOM As Boolean = True)

    Dim iFileNo                     As Integer

    iFileNo = FreeFile
    OpenForOutput the_sFilename, iFileNo, the_nCodePage, the_bPrefixWithBOM
    Print_ iFileNo, the_sFileContents, the_nCodePage, False
    Close iFileNo

End Sub

