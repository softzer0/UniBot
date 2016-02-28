Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function SetCurrentDirectoryA Lib "kernel32" (ByVal lpPathName As String) As Long

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

Private Declare Function EnumThreadWindows Lib "user32" _
   (ByVal dwThreadId As Long, ByVal lpfn As Long, _
   ByVal lParam As Long) As Long
Private Declare Function GetWindowLongA Lib "user32" _
   (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
   (Destination As Any, Source As Any, ByVal Length As Long)

Private Const GWL_HWNDPARENT As Long = -8&

Private Declare Function GetClassname Lib "user32" _
   Alias "GetClassNameA" _
   (ByVal hWnd As Long, ByVal lpClassName As String, _
   ByVal nMaxCount As Long) As Long

Public Function Classname(ByVal hWnd As Long) As String
   Dim nRet As Long
   Dim Class As String
   Const MaxLen As Long = 256
   
   ' Retrieve classname of passed window.
   Class = String$(MaxLen, 0)
   nRet = GetClassname(hWnd, Class, MaxLen)
   If nRet Then Classname = Left$(Class, nRet)
End Function

Private Function EnumThreadWndProc(ByVal hWnd As Long, _
   ByVal lpResult As Long) As Long
   Dim nStyle As Long
   Dim Class As String
   
   ' Assume we will continue enumeration.
   EnumThreadWndProc = True
   
   ' Test to see if this window is parented.
   ' If not, it may be what we're looking for!
   If GetWindowLongA(hWnd, GWL_HWNDPARENT) = 0 Then
      ' This rules out IDE windows when not compiled.
      Class = Classname(hWnd)
      ' Version agnostic test.
      If InStr(Class, "Thunder") = 1 Then
         If InStr(Class, "Main") = (Len(Class) - 3) Then
            ' Copy hWnd to result variable pointer,
            Call CopyMemory(ByVal lpResult, hWnd, 4&)
            ' and stop enumeration.
            EnumThreadWndProc = False
         End If
      End If
   End If
End Function

Public Function FindHiddenTopWindow() As Long
   ' This function returns the hidden toplevel window
   ' associated with the current thread of execution.
   Call EnumThreadWindows(App.ThreadID, _
      AddressOf EnumThreadWndProc, VarPtr(FindHiddenTopWindow))
End Function

Function CommDlg(Optional bolSave As Boolean, Optional strDialogTitle As String, Optional strFilter As String = "Any file|*.*", Optional strInitDir As String, Optional strDefFile As String, Optional lngFlags As FileOpenConstants) As String
CD.hWndOwner = FindHiddenTopWindow
CD.hInstance = App.hInstance
If strDialogTitle = vbNullString Then If bolSave Then strDialogTitle = "Save" Else: strDialogTitle = "Open"
CD.lpstrTitle = strDialogTitle
CD.lpstrFilter = Replace(strFilter, "|", Chr$(0)) + Chr$(0)
If strInitDir <> vbNullString And strInitDir <> vbNullChar Then CD.lpstrInitialDir = strInitDir Else: CD.lpstrInitialDir = CurDir$ 'If strInitDir = "1" Then CD.lpstrInitialDir = strInitD Else:
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

Private Sub Main()
On Error GoTo E
Dim strP As String: strP = Replace(Command$, """", vbNullString)
If strP <> vbNullString Then If Dir$(strP) = vbNullString Then strP = vbNullString
If strP = vbNullString Then strP = CommDlg(, "Select UniBot plugin project", "VB6 project (*.vbp)|*.vbp")
If strP = vbNullString Then Exit Sub
Open strP For Input As #1
Dim strT(1) As String, bytT As Byte, bolT As Boolean, strC() As String, strT1(1) As String
ReDim strC(0)
While Not EOF(1)
Line Input #1, strT(0)
strT(0) = Trim$(strT(0))
Select Case True
Case StrComp(Left$(strT(0), 6), "Class=", vbTextCompare) = 0, StrComp(Left$(strT(0), 5), "Form=", vbTextCompare) = 0
If strC(0) <> vbNullString Then ReDim Preserve strC(UBound(strC) + 1)
strC(UBound(strC)) = strT(0)
strT(0) = vbNullString
Case StrComp(Left$(strT(0), 10), "ExeName32=", vbTextCompare) = 0: strT(0) = Split(Left$(strT(0), Len(strT(0)) - 1), ".")(0) & ".\EXT\"""
Case StrComp(Left$(strT(0), 8), "Startup=", vbTextCompare) = 0: strT1(0) = Mid$(strT(0), 10, Len(strT(0)) - 10): strT(0) = vbNullString
Case StrComp(Left$(strT(0), 15), "ThreadingModel=", vbTextCompare) = 0, StrComp(Left$(strT(0), 10), "StartMode=", vbTextCompare) = 0: strT(0) = vbNullString
Case StrComp(Left$(strT(0), 10), "Reference=", vbTextCompare) = 0: If Right$(strT(0), 23) = "#UniBot_PluginInterface" Then bytT = bytT + 1
Case StrComp(strT(0), "Type=OleDll", vbTextCompare) = 0
strT(0) = "Type=Exe"
bolT = True
bytT = bytT + 1
Case StrComp(strT(0), "Type=Exe", vbTextCompare) = 0
strT(0) = "Type=OleDll"
bytT = bytT + 1
End Select
If strT(0) <> vbNullString Then strT(1) = strT(1) & strT(0) & vbCrLf
Wend
Close #1
If bytT < 2 Then MsgBox "This isn't UniBot plugin project!", vbCritical: Exit Sub
If bolT Then
strT(1) = Replace(strT(1), "\EXT\", "exe", , 1)
SetCurrentDirectoryA Left$(strP, InStrRev(strP, "\"))
Dim bytT1(1) As Byte
strT1(0) = vbNullString
For bytT = 0 To UBound(strC)
If StrComp(Left$(strC(bytT), 6), "Class=", vbTextCompare) = 0 Then
Open Trim$(Mid$(strC(bytT), InStr(strC(bytT), ";") + 1)) For Input Access Read As #1
While Not EOF(1)
Line Input #1, strT(0)
strT(0) = Trim$(strT(0))
Select Case True
Case StrComp(Left$(strT(0), 7), "Private", vbTextCompare) = 0 And bytT1(1) = 0, StrComp(Left$(strT(0), 6), "Public", vbTextCompare) = 0 And bytT1(1) = 0: bytT1(1) = 1
Case StrComp(Left$(strT(0), 3), "Dim", vbTextCompare) = 0, StrComp(Left$(strT(0), 7), "Declare", vbTextCompare) = 0, StrComp(Left$(strT(0), 5), "Const", vbTextCompare) = 0, StrComp(Left$(strT(0), 4), "Type", vbTextCompare) = 0, StrComp(Left$(strT(0), 4), "Enum", vbTextCompare) = 0, StrComp(Left$(strT(0), 6), "Global", vbTextCompare) = 0, StrComp(Left$(strT(0), 6), "Option", vbTextCompare) = 0, StrComp(Left$(strT(0), 10), "Implements", vbTextCompare) = 0, StrComp(Left$(strT(0), 3), "#If", vbTextCompare) = 0, StrComp(Left$(strT(0), 6), "#Const", vbTextCompare) = 0, StrComp(Left$(strT(0), 3), "Def", vbTextCompare) = 0: If bytT1(1) = 0 Then bytT1(1) = 1
Case StrComp(Left$(strT(0), 3), "Sub", vbTextCompare) = 0, StrComp(Left$(strT(0), 11), "Private Sub", vbTextCompare) = 0, StrComp(Left$(strT(0), 10), "Public Sub", vbTextCompare) = 0, StrComp(Left$(strT(0), 8), "Function", vbTextCompare) = 0, StrComp(Left$(strT(0), 16), "Private Function", vbTextCompare) = 0, StrComp(Left$(strT(0), 15), "Public Function", vbTextCompare) = 0: If bytT1(1) = 1 Then bytT1(1) = 2
Case StrComp(Left$(strT(0), 3), "End", vbTextCompare) <> 0 And bytT1(1) = 2: strT(0) = Replace(strT(0), "DoEvents = 1", "DoEvents = 2")
End Select
If bytT1(1) > 0 Then strT1(bytT1(1) - 1) = strT1(bytT1(1) - 1) & strT(0) & vbCrLf
Select Case True
Case StrComp(Left$(strT(0), 13), "MultiUse = -1", vbTextCompare) = 0, StrComp(Left$(strT(0), 29), "Attribute VB_Creatable = True", vbTextCompare) = 0, StrComp(Left$(strT(0), 27), "Attribute VB_Exposed = True", vbTextCompare) = 0: bytT1(0) = bytT1(0) + 1
Case StrComp(Left$(strT(0), 27), "Implements IPluginInterface", vbTextCompare) = 0
bytT1(0) = bytT1(0) + 1
strT1(0) = "'" & strC(bytT) & vbCrLf & strT1(0)
Case bolT: If StrComp(Left$(strT(0), 29), "Private Sub Class_Terminate()", vbTextCompare) = 0 Then bolT = False
End Select
Wend
Close #1
If bytT1(0) <> 4 Then
bytT1(0) = 0
strT1(0) = vbNullString
bolT = True
Else: Exit For
End If
End If
Next
If bytT1(0) = 4 Then
strP = "TEST_" & Mid$(strP, InStrRev(strP, "\") + 1)
strT(0) = vbNullString
If Dir$(strP) = vbNullString Then PrepName strP, strT(0), "TESTING?.frm", " (?)"
strT(1) = strT(1) & "Form=TESTING" & strT(0) & ".frm" & vbCrLf & "Startup=""frmTesting"""
Open strP For Output Access Write As #1
Print #1, strT(1)
PrintOther strC, bytT
Close #1
Open "TESTING" & strT(0) & ".frm" For Output Access Write As #1
strT1(0) = Replace(StrConv(LoadResData(101, "TEXT"), vbUnicode), "DOREPLACEHERE", Left$(strT1(0), Len(strT1(0)) - 2), , 1)
If bolT Then strT1(0) = Replace(strT1(0), "Class_Terminate" & vbCrLf, vbNullString, , 1)
Print #1, strT1(0)
Print #1, Left$(strT1(1), Len(strT1(1)) - 2)
Close #1
Else
E:
MsgBox "This project doesn't have any form/class which fullfills the requirements to be a plugin!", vbCritical
Exit Sub
End If
Else
If strT1(0) = vbNullString Then MsgBox "This project doesn't have startup form set!", vbExclamation: Exit Sub
strT(1) = Replace(strT(1), "\EXT\", "dll", , 1)
SetCurrentDirectoryA Left$(strP, InStrRev(strP, "\"))
Dim strT2 As String: strT2 = strT1(0)
strT1(0) = "VERSION 1.0 CLASS" & vbCrLf & "BEGIN" & vbCrLf & "  MultiUse = -1  'True" & vbCrLf & "  Persistable = 0  'NotPersistable" & vbCrLf & "  DataBindingBehavior = 0  'vbNone" & vbCrLf & "  DataSourceBehavior  = 0  'vbNone" & vbCrLf & "  MTSTransactionMode  = 0  'NotAnMTSObject" & vbCrLf & "END" & vbCrLf
For bytT = 0 To UBound(strC)
If StrComp(Left$(strC(bytT), 5), "Form=", vbTextCompare) = 0 Then
Open RTrim$(Mid$(strC(bytT), 6)) For Input As #1
Line Input #1, strT(0)
If EOF(1) Then GoTo E
Line Input #1, strT(0)
If StrComp(Left$(strT(0), Len(strT2) + 14), "Begin VB.Form " & strT2, vbTextCompare) = 0 Then
While Not EOF(1)
Line Input #1, strT(0)
strT(0) = Trim$(strT(0))
If StrComp(Left$(strT(0), 10), "Attribute ", vbTextCompare) = 0 Then
bolT = True
If StrComp(Left$(strT(0), 21), "Attribute VB_NAME = """, vbTextCompare) = 0 Then
strT(0) = "Attribute VB_NAME = ""DOREPLACEHERE"""
ElseIf StrComp(Left$(strT(0), 30), "Attribute VB_Creatable = False", vbTextCompare) = 0 Then strT(0) = "Attribute VB_Creatable = True"
ElseIf StrComp(Left$(strT(0), 33), "Attribute VB_PredeclaredId = True", vbTextCompare) = 0 Then strT(0) = "Attribute VB_PredeclaredId = False"
ElseIf StrComp(Left$(strT(0), 28), "Attribute VB_Exposed = False", vbTextCompare) = 0 Then strT(0) = "Attribute VB_Exposed = True"
End If
ElseIf Left$(strT(0), 1) = "'" Then
If Trim$(Replace(Mid$(strT(0), 2), "-", vbNullString)) = "Added" Then
bolT = Not bolT
GoTo N
ElseIf Left$(strT(0), 7) = "'Class=" Then
strT(0) = Mid$(strT(0), 8)
strT2 = Split(strT(0), ";")(0)
strT1(1) = LTrim$(Mid$(strT(0), InStr(strT(0), ";") + 1))
GoTo N
End If
ElseIf bolT Then strT(0) = Replace(strT(0), "DoEvents = 2", "DoEvents = 1")
End If
If bolT Then strT1(0) = strT1(0) & strT(0) & vbCrLf
N:
Wend
Close #1
If strT2 = vbNullString Then
strT2 = "Class1"
strT1(1) = strT2
End If
Dim strT3 As String: strT3 = Left$(strP, InStrRev(strP, "\"))
strP = Mid$(strP, Len(strT3) + 1)
If InStr(Left$(strP, InStr(strP & ".", ".") - 1), "_") > 0 And InStr(Left$(strP, InStr(strP & ".", ".") - 1), "_") < Len(strP) - Len(Mid$(strP, InStr(strP & ".", "."))) Then strP = Mid$(strP, InStr(strP, "_") + 1) Else: strP = "p_" & strP
strP = strT3 & strP
strT1(0) = Replace(strT1(0), "DOREPLACEHERE", strT2, , 1)
strT1(1) = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(strT1(1), "*", vbNullString), "/", vbNullString), ":", vbNullString), "<", vbNullString), ">", vbNullString), "?", vbNullString), "\", vbNullString), "|", vbNullString)
If InStr(strT1(1), ".") = 0 Then strT1(1) = strT1(1) & ".cls"
strT(0) = vbNullString
If Dir$(strP) = vbNullString Then
strT1(1) = Left$(strT1(1), InStr(strT1(1), ".") - 1) & "?" & Mid$(strT1(1), InStr(strT1(1), "."))
PrepName strP, strT(0), strT1(1), "_?"
strT1(1) = Replace(strT1(1), "?", strT(0), , 1)
End If
Open strP For Output Access Write As #1
Print #1, Left$(strT(1), Len(strT(1)) - 2)
Print #1, "Startup=""(None)"""
Print #1, "Class=" & strT2 & "; " & strT1(1)
PrintOther strC, bytT
Close #1
Open strT1(1) For Output Access Write As #1
Print #1, Left$(strT1(0), Len(strT1(0)) - 2)
Close #1
Exit For
End If
Close #1
End If
Next
If Not bolT Then GoTo E
End If
If MsgBox("All is done! Do you want to open the project now?", vbInformation + vbYesNo) = vbYes Then Shell "cmd.exe /c """ & strP & """", vbHide
End Sub

Private Sub PrintOther(strC() As String, bytT As Byte)
Dim intT As Integer
For intT = 0 To bytT - 1
Print #1, strC(intT)
Next
For intT = bytT + 1 To UBound(strC)
If strC(bytT) <> vbNullString Then Print #1, strC(intT)
Next
End Sub

Private Sub PrepName(strP As String, strT0 As String, strT1 As String, strT2 As String)
If Dir$(Replace(strT1, "?", vbNullString, , 1)) = vbNullString Then Exit Sub
Dim bytT As Byte
Do
bytT = bytT + 1
strT0 = Replace(strT2, "?", bytT, , 1)
Loop Until Dir$(Replace(strT1, "?", strT0, , 1)) = vbNullString
End Sub
