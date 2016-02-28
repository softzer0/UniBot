Attribute VB_Name = "modBrowse"
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_EDITBOX = &H10
Private Const BIF_VALIDATE = &H20
Private Const BIF_NEWDIALOGSTYLE = &H40
Private Const BIF_USENEWUI = (BIF_NEWDIALOGSTYLE Or BIF_EDITBOX)
Private Const BIF_BROWSEINCLUDEURLS = &H80
Private Const BIF_UAHINT = &H100
Private Const BIF_NONEWFOLDERBUTTON = &H200
Private Const BIF_NOTRANSLATETARGETS = &H400
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000
Private Const BIF_BROWSEINCLUDEFILES = &H4000
Private Const BIF_SHAREABLE = &H8000
Private Const MAX_PATH = 260
Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_SETSELECTION = (WM_USER + 102)

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BrowseInfo) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Private mstrSTARTFOLDER As String
Function GetFolder(ByVal hWndModal As Long, Optional StartFolder As String, Optional Title As String = "Please select a folder:", _
    Optional IncludeFiles As Boolean = False, Optional IncludeNewFolderButton As Boolean = False) As String
     Dim bInf As BrowseInfo
     Dim RetVal As Long
     Dim PathID As Long
     Dim RetPath As String
     Dim Offset As Integer
     'Set the properties of the folder dialog
     bInf.hWndOwner = hWndModal
     bInf.pIDLRoot = 0
     bInf.lpszTitle = Title
     bInf.ulFlags = BIF_RETURNONLYFSDIRS Or BIF_STATUSTEXT
     If IncludeFiles Then bInf.ulFlags = bInf.ulFlags Or BIF_BROWSEINCLUDEFILES
     If IncludeNewFolderButton Then bInf.ulFlags = bInf.ulFlags Or BIF_NEWDIALOGSTYLE
     If StartFolder = vbNullChar Then StartFolder = strInitD
     If StartFolder <> vbNullString Then
        mstrSTARTFOLDER = StartFolder & vbNullChar
        bInf.lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc) 'get address of function.
    End If
     'Show the Browse For Folder dialog
     PathID = SHBrowseForFolder(bInf)
     RetPath = Space$(512)
     RetVal = SHGetPathFromIDList(ByVal PathID, ByVal RetPath)
     If RetVal Then
          'Trim off the null chars ending the path
          'and display the returned folder
          Offset = InStr(RetPath, Chr$(0))
          GetFolder = Left$(RetPath, Offset - 1)
          'Free memory allocated for PIDL
          CoTaskMemFree PathID
     Else
          GetFolder = vbNullChar
     End If
End Function
Private Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
    On Error Resume Next
    Dim lpIDList As Long
    Dim ret As Long
    Dim sBuffer As String
    Select Case uMsg
        Case BFFM_INITIALIZED
            Call SendMessage(hWnd, BFFM_SETSELECTION, 1, mstrSTARTFOLDER)
        Case BFFM_SELCHANGED
            sBuffer = Space(MAX_PATH)
            ret = SHGetPathFromIDList(lp, sBuffer)
            If ret = 1 Then
                Call SendMessage(hWnd, BFFM_SETSTATUSTEXT, 0, sBuffer)
            End If
    End Select
    BrowseCallbackProc = 0
End Function
Private Function GetAddressofFunction(add As Long) As Long
  GetAddressofFunction = add
End Function


