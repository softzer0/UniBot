Attribute VB_Name = "modStuff"
Option Explicit

Public Plugins As New Collection, strInitD As String

Public Const strDigit = "0123456789"
Public Const strLett = "qwertyuiopasdfghjklzxcvbnm"
Public Const strULett = "QWERTYUIOPASDFGHJKLZXCVBNM"
Public Const strSym = "~`!@#$%^&*()-=_+[]\{}|;':"",./<>?"

Dim Hash As New MD5Hash

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Const LB_SETHORIZONTALEXTENT = &H194

Public Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" ( _
      ByVal hInst As Long, _
      ByVal lpsz As String, _
      ByVal uType As Long, _
      ByVal cxDesired As Long, _
      ByVal cyDesired As Long, _
      ByVal fuLoad As Long _
   ) As Long

Public Const IMAGE_ICON = 1

Public Const CP_UTF8       As Long = 65001      ' UTF8.
Public Const CP_ACP        As Long = 0          ' Default ANSI code page.
Public Const CP_UTF16_LE   As Long = 1200       ' UTF16 - little endian.

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

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) _
        As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
        (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As _
        Long) As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
        (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Public Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal _
        dest As Long, ByVal src As Long, ByVal Length As Long) As Long

Private Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

Public Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

   Public Declare Function SendMessage Lib "user32" _
      Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
      ByVal wParam As Long, lParam As Any) As Long

   Private Const WM_SETTEXT = &HC
   Private Const EM_SETSEL = &HB1

'**************************************
'Windows API/Global Declarations for :Common Dialog without OCX
'**************************************
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private CD As OPENFILENAME
Private Type OPENFILENAME
lStructSize As Long
hWndOwner As Long
hInstance As Long
lpstrFilter As String
lpstrCustomFilter As String
nMaxCustFilter As Long
nFilterIndex As Long
lpstrFile As String
nMaxFile As Long
lpstrFileTitle As String
nMaxFileTitle As Long
lpstrInitialDir As String
lpstrTitle As String
FLAGS As Long
nFileOffset As Integer
nFileExtension As Integer
lpstrDefExt As String
lCustData As Long
lpfnHook As Long
lpTemplateName As String
End Type
Enum FileOpenConstants
cdlOFNOverwritePrompt = 2
cdlOFNHideReadOnly = 4
cdlOFNPathMustExist = 2048
cdlOFNFileMustExist = 4096
cdlOFNNoReadOnlyReturn = 32768
cdlOFNExplorer = 524288
End Enum

Public Sub SetTopMostWindow(hWnd As Long, Topmost As Boolean)
  If Topmost Then 'Make the window topmost
     SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, _
        0, SWP_NOMOVE Or SWP_NOSIZE
  Else
     SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, _
        0, 0, SWP_NOMOVE Or SWP_NOSIZE
  End If
End Sub

Function TrimComma(strT As String) As String
TrimComma = strT
Do While InStr(TrimComma, ",,") > 0
TrimComma = Replace(TrimComma, ",,", ",")
Loop
TrimComma = Replace(Replace(Replace(Replace(TrimComma, "(", vbNullString), ")", vbNullString), "+", vbNullString), "'", vbNullString)
End Function

Sub LoadLst(lst As ListBox)
If frmPlugins.strRg <> vbNullString Then lst.AddItem "AdvancedRegEx"
Dim s() As String, i As Integer
s() = Split(frmPlugins.strPl, vbLf)
For i = 0 To UBound(s) - 1
lst.AddItem s(i)
Next
End Sub

Sub SetListboxScrollbar1(ByVal lst As ListBox)
Dim i As Integer
Dim new_len As Long
Dim max_len As Long

    For i = 0 To lst.ListCount - 1
        new_len = 10 + lst.Parent.ScaleX( _
            lst.Parent.TextWidth(lst.list(i)), _
            lst.Parent.ScaleMode, vbPixels)
        If max_len < new_len Then max_len = new_len
    Next i

    SendMessage lst.hWnd, _
        LB_SETHORIZONTALEXTENT, _
        max_len, 0
End Sub

Function CreateRG() As Object
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

' // Load all plugins
Sub ScanPlugins(strFolder As String, Optional bolSubFolder As Boolean)
    Dim fso As FileSystemObject
    Dim fld As folder
    
    Set fso = New FileSystemObject
    
    ' Scan plugins
    Scan_ fso.GetFolder(strFolder), bolSubFolder
    
End Sub

' // Scan folder (also scan sub-folders) and loading plugins
Private Sub Scan_(fld As folder, Optional subf As Boolean)
On Error Resume Next
    Dim subFld  As folder
    Dim fle     As File
    Dim Index   As Long
    Dim strT As String
    
    strT = Replace(frmPlugins.strP, "|", "\")
    
    ' Check all files in folder
    For Each fle In fld.Files()
        ' Get extension pos in name
        Index = InStrRev(fle.Name, ".")
        
        If Index Then
            ' If extension is "dll"
            If StrComp(Mid$(fle.Name, Index + 1), "dll", vbTextCompare) = 0 Then
            
            If StrComp(fle.Name, "dotnetcomregexlib.dll", vbTextCompare) = 0 Then GoTo N
            
              If InStr(strT, fle.Name) = 0 Then frmPlugins.strNL = frmPlugins.strNL & Replace(fle.path, frmPlugins.strLocation, vbNullString) & vbLf
            
            End If
            
        End If
N:
    Next
    
    If subf Then
      ' Check sub-folders
      For Each subFld In fld.SubFolders()
          Scan_ subFld ', True
      Next
    End If
    
End Sub

'**************************************
' Name: Common Dialog without OCX
' Description:Hi All,
' The Perpose of this Progarm is to Use windows Common Dialog Control Control Without the COMDLG32.OCX file. This will work even if the File is not Present
' This is only for Open and Save Functions. But You can append it to get Color and other Dialog Boxes too,
' Just Send any comments to
' visual_basic@ manjulapra.com
' Visit me at
' http://www.manjulapra.com
' Thank You
' By: Manjula Dharmawardhana
' Modified by: MikiSoft
'
' Inputs:The Filter for the Common Dialog
' The Default Extention for the Common Dialog
' Optionally the Dialog Titile
'
' Returns:The Path of the Selected File
'
' Side Effects:None Identified
'
'This code is copyrighted and has' limited warranties.Please see http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=13368&lngWId=1'for details.
'**************************************

Function CommDlg(Optional bolSave As Boolean, Optional strDialogTitle As String, Optional strFilter As String = "Any file|*.*", Optional strInitDir As String, Optional strDefFile As String, Optional lngFlags As FileOpenConstants) As String
CD.hWndOwner = Screen.ActiveForm.hWnd
CD.hInstance = App.hInstance
If strDialogTitle = vbNullString Then If bolSave Then strDialogTitle = "Save" Else: strDialogTitle = "Open"
CD.lpstrTitle = strDialogTitle
CD.lpstrFilter = Replace(strFilter, "|", Chr$(0)) + Chr$(0)
CD.lpstrDefExt = "*.*"
If strInitDir <> vbNullString And strInitDir <> vbNullChar Then CD.lpstrInitialDir = strInitDir Else: CD.lpstrInitialDir = CurDir$ 'If strInitDir = "1" Then CD.lpstrInitialDir = strInitD Else:
CD.lpstrFile = strDefFile & Chr$(0) & Space$(259 - Len(strDefFile))
CD.nMaxFile = 260
CD.lStructSize = Len(CD)
If bolSave Then
If lngFlags = 0 Then CD.FLAGS = cdlOFNExplorer + cdlOFNNoReadOnlyReturn + cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNOverwritePrompt & Chr$(0) Else: CD.FLAGS = lngFlags & Chr$(0)
If GetSaveFileName(CD) = 1 Then CommDlg = CD.lpstrFile
Else
If lngFlags = 0 Then CD.FLAGS = cdlOFNExplorer + cdlOFNPathMustExist + cdlOFNFileMustExist & Chr$(0) Else: CD.FLAGS = lngFlags & Chr$(0)
If GetOpenFileName(CD) = 1 Then CommDlg = CD.lpstrFile
End If
Dim pos As Integer: pos = InStr(CommDlg, Chr$(0))
If pos > 0 Then CommDlg = Left$(CommDlg, pos - 1)
End Function

Function LoadBig(txt As TextBox, FileOrString As String, Optional IsString As Boolean, Optional Pre As String)
    On Error GoTo E
   
    Dim TempText As String
    Dim iret As Long
    If IsString Then TempText = Pre & FileOrString Else: TempText = Pre & LoadFile(FileOrString)
    txt.Text = vbNullString
    
    iret = SendMessage(txt.hWnd, WM_SETTEXT, 0&, ByVal TempText)
    'iret = SendMessage(txt.hWnd, WM_GETTEXTLENGTH, 0&, ByVal 0&)
    'Debug.Print "WM_GETTEXTLENGTH: " & iret
E:
End Function

Function LoadFile(strPath As String) As String
On Error GoTo E
If Dir$(strPath, vbHidden) = vbNullString Then Exit Function
LoadFile = String$(FileLen(strPath), vbNullChar)
Open strPath For Binary Access Read As #3
Get #3, , LoadFile
Close #3
E:
End Function

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

Function SelectT(txt As Object, SelStart As Long, SelEnd As Long)
    Dim res As Long
    res = SendMessage(txt.hWnd, EM_SETSEL, SelStart, SelEnd)
End Function

Function RegExpr(myPattern As String, myString As String, Optional myReplace As String, Optional bytResults As Byte, Optional intStart As Integer, Optional intCount As Integer) As Variant
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

Function CheckText(obj As TextBox)
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

Function StrAR(i As Byte, Optional bolR As Boolean, Optional bolS As Boolean) As Byte
Dim intT As Integer
If Not bolR Then
If InStr(frmMain.cmdOpt(0).Tag, "-" & i & "-") > 0 Then
intT = Split(Split(frmMain.cmdOpt(0).Tag, "-" & i & "-")(1), vbLf)(0)
frmMain.cmdOpt(0).Tag = Replace(frmMain.cmdOpt(0).Tag, "-" & i & "-" & intT, "-" & i & "-" & intT + 1, , 1)
Else
frmMain.cmdOpt(0).Tag = frmMain.cmdOpt(0).Tag & "-" & i & "-1" & vbLf
If Not bolS Then If frmMain.cmbIndex.ListCount - 1 = frmMain.cmbIndex.ListIndex Then frmMain.AddI
End If
Else
intT = Split(Split(frmMain.cmdOpt(0).Tag, "-" & i & "-")(1), vbLf)(0) - 1
If Not bolS And intT = 0 Then
If Not frmMain.Filled(i, 3) Then
If frmMain.RemI(True) Then
StrAR = 1
Exit Function
Else: StrAR = 2
End If
End If
End If
frmMain.cmdOpt(0).Tag = Replace(frmMain.cmdOpt(0).Tag, "-" & i & "-" & intT + 1 & vbLf, IIf(intT > 0, "-" & i & "-" & intT & vbLf, vbNullString), , 1)
End If
'Debug.Print frmMain.cmdOpt(0).Tag
End Function

Function CheckPublic(ByVal strN As String, Optional bolP As Boolean, Optional bolS As Boolean) As Boolean
Dim bytT As Byte
If InStr(frmMain.fraS.Tag, vbLf & strN & vbLf) > 0 Then bytT = Split(Split(frmMain.fraS.Tag, vbLf & strN & vbLf)(1), vbLf)(0)
If bolP Then
If InStr(frmMain.fraS.Tag, vbLf & strN & vbLf) > 0 Then
If Not bolS Then
If MsgBox("Public string with same name already exists. Keep changes?", vbQuestion + vbYesNo) = vbNo Then
CheckPublic = True
Exit Function
End If
End If
frmMain.fraS.Tag = Replace(frmMain.fraS.Tag, vbLf & strN & vbLf & bytT & vbLf, vbLf & strN & vbLf & bytT + 1 & vbLf, , 1)
Else: frmMain.fraS.Tag = vbLf & strN & vbLf & "0" & frmMain.fraS.Tag
End If
ElseIf bytT > 0 Then
frmMain.fraS.Tag = Replace(frmMain.fraS.Tag, vbLf & strN & vbLf & bytT & vbLf, vbLf & strN & vbLf & bytT - 1 & vbLf, , 1)
Else: frmMain.fraS.Tag = Replace(frmMain.fraS.Tag, vbLf & strN & vbLf & "0" & vbLf, vbLf, , 1)
End If
'Debug.Print frmMain.fraS.Tag
End Function

Function ProcessNumber(strT As Variant, Optional bolI As Boolean, Optional bolA As Boolean) As Integer
Dim intT As Integer
If Not bolI Then intT = 255 Else: intT = 32767
If strT <= intT Then
If strT < 0 Then ProcessNumber = strT * (-1) Else: ProcessNumber = strT
Else: If Not bolA Then ProcessNumber = intT Else: ProcessNumber = 0
End If
End Function

Function ReplaceString(ByVal strInp As String, Optional strSrc As String = vbNullChar) As String
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

Function AddChrs(strR As String, strT As String, Optional strInp As String, Optional intL As Integer, Optional intM As Variant) As String
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
ElseIf InStr(frmPlugins.strC, "," & Left$(strT, InStr(strT, "(") - 1) & ",") > 0 Then
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
strT2(0) = Split(frmPlugins.strC, "," & strComm(0) & ",")(0)
strT2(0) = Mid$(strT2(0), InStrRev(vbLf & strT2(0), vbLf, Len(strT2(0)) - 1) + 1)
Set tmpObj = Plugins.Item(Left$(strT2(0), InStr(strT2(0) & "|", "|") - 1))
frmMain.bytPlgUse = frmMain.bytPlgUse + 1
strT = "'" & Replace(tmpObj.Execute(strComm), "'", "''") & "'"
frmMain.bytPlgUse = frmMain.bytPlgUse - 1
If frmMain.bolUnl And frmMain.bytPlgUse = 0 Then Unload frmMain: Exit Function
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

Function FindC1(strS As String, Optional intC As Long = 2, Optional strC As String = "'") As Long
FindC1 = InStr(intC, Left$(strS, intC - 1) & Replace(Mid$(strS, intC), strC & strC, "  "), strC)
End Function

Function FindSep(strExp As String, Optional intS As Long = 1, Optional strC As String = ",", Optional strE As String = "'", Optional comp As Integer = vbBinaryCompare) As Long
Dim intC(2) As Long
intC(1) = intS
intC(2) = intS
Do
intC(0) = intC(2)
intC(1) = InStr(intC(2), strExp, strE) + 1
If intC(1) = 1 Then Exit Do
intC(2) = FindC1(strExp, intC(1), strE) + 1
Loop Until InStr(1, Mid$(strExp, intC(0), intC(1) - intC(0) - 1), strC, comp) > 0 Or intC(2) = 1
If intC(2) = 1 Then FindSep = InStr(intC(1), strExp, strC, comp) Else: FindSep = InStr(intC(0), strExp, strC, comp)
End Function

Private Function RandNum(ByVal Low As Long, ByVal High As Long) As Long
RandNum = Int((High - Low + 1) * Rnd) + Low
End Function

Function RandStr(Optional strR As String, Optional intL As Integer) As String
    If strR = vbNullString Then strR = strDigit & strLett & strULett & strSym
    If intL = 0 Then intL = 15
    Dim i As Integer
    For i = 1 To intL
        RandStr = RandStr & Mid$(strR, Int(Rnd() * Len(strR) + 1), 1)
    Next
End Function

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
