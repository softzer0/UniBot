Attribute VB_Name = "modS"
Option Explicit

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
        (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As _
        Long) As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
        (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Public Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal _
        dest As Long, ByVal src As Long, ByVal Length As Long) As Long

Public Declare Function SendMessage Lib "user32" _
      Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
      ByVal wParam As Long, lParam As Any) As Long
   Private Const WM_SETTEXT = &HC
   Private Const EM_SETSEL = &HB1
   
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40

Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

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
flags As Long
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
CD.hWndOwner = Screen.ActiveForm.hwnd
CD.hInstance = App.hInstance
If strDialogTitle = vbNullString Then If bolSave Then strDialogTitle = "Save" Else: strDialogTitle = "Open"
CD.lpstrTitle = strDialogTitle
CD.lpstrFilter = Replace(strFilter, "|", Chr$(0)) + Chr$(0)
CD.lpstrDefExt = "*.*"
If strInitDir <> vbNullString And strInitDir <> vbNullChar Then CD.lpstrInitialDir = strInitDir Else: CD.lpstrInitialDir = CurDir$ 'If strInitDir = "1" Then CD.lpstrInitialDir = App.Path Else:
CD.lpstrFile = strDefFile & Chr$(0) & Space$(259 - Len(strDefFile))
CD.nMaxFile = 260
CD.lStructSize = Len(CD)
If bolSave Then
If lngFlags = 0 Then CD.flags = cdlOFNExplorer + cdlOFNNoReadOnlyReturn + cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNOverwritePrompt & Chr$(0) Else: CD.flags = lngFlags & Chr$(0)
If GetSaveFileName(CD) = 1 Then CommDlg = CD.lpstrFile
Else
If lngFlags = 0 Then CD.flags = cdlOFNExplorer + cdlOFNPathMustExist + cdlOFNFileMustExist & Chr$(0) Else: CD.flags = lngFlags & Chr$(0)
If GetOpenFileName(CD) = 1 Then CommDlg = CD.lpstrFile
End If
Dim pos As Integer: pos = InStr(CommDlg, Chr$(0))
If pos > 0 Then CommDlg = Left$(CommDlg, pos - 1)
End Function

Public Sub SetTopMostWindow(hwnd As Long, Topmost As Boolean)
  If Topmost Then 'Make the window topmost
     SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, _
        0, SWP_NOMOVE Or SWP_NOSIZE
  Else
     SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, _
        0, 0, SWP_NOMOVE Or SWP_NOSIZE
  End If
End Sub

Function LoadBig(txt As TextBox, FileOrString As String, Optional IsString As Boolean, Optional Pre As String)
    On Error GoTo E
   
    Dim TempText As String
    Dim iret As Long
    If IsString Then TempText = Pre & FileOrString Else: TempText = Pre & LoadFile(FileOrString)
    'DoEvents
    txt.Text = vbNullString
    
    iret = SendMessage(txt.hwnd, WM_SETTEXT, 0&, ByVal TempText)
    'iret = SendMessage(txt.hWnd, WM_GETTEXTLENGTH, 0&, ByVal 0&)
    'Debug.Print "WM_GETTEXTLENGTH: " & iret
E:
End Function

Function LoadFile(strPath As String) As String
On Error GoTo E
If Dir$(strPath, vbHidden) = vbNullString Then Exit Function
LoadFile = String$(FileLen(strPath), vbNullChar)
Open strPath For Binary Access Read As #2
Get #2, , LoadFile
Close #2
E:
End Function

Function SelectT(txt As Object, SelStart As Long, SelEnd As Long)
    Dim res As Long
    res = SendMessage(txt.hwnd, EM_SETSEL, SelStart, SelEnd)
End Function
