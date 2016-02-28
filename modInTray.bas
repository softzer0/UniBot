Attribute VB_Name = "modInTray"
'----------------------------------------------------------------------
'Start of Module: modInTray

Option Explicit

'Windows API's
Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" _
        (Class As WNDCLASS) As Long
Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" _
        (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal _
        lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, _
        ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
        ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance _
        As Long, lpParam As Any) As Long
Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias _
        "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As _
        NOTIFYICONDATA) As Long
Private Declare Function RegisterWindowMessage Lib "user32" Alias _
        "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" _
        (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
        ByVal lParam As Long) As Long
Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" _
        (ByVal lpClassName As String, ByVal hInstance As Long) As Long

Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Type WNDCLASS
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As Long
    lpszClassName As String
End Type

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_USER = &H400
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_USER_SYSTRAY = WM_USER + 5
Public Const WM_CLOSE = &H10
Public Const GWL_USERDATA = (-21)

'Have we initialized TaskbarRestart yet?
Private m_bMessageInited As Boolean
'Windows message sent out when the taskbar is restarted
' (ie4 and up)
Private m_uTaskbarRestart As Long

'Dummy function to allow AddressOf to assign to a variable
Public Function Pass(N As Long) As Long
    Pass = N
End Function

'Return a VB object pointed by nPointer
Public Function DeRef(nPointer As Long) As clsInTray
    CopyMemory VarPtr(DeRef), VarPtr(nPointer), 4
End Function

'Creates a pointer to a VB object
Public Function CreateRef(obj As clsInTray) As Long
    CopyMemory VarPtr(CreateRef), VarPtr(obj), 4
End Function

'Destroys a VB object created by DeRef (otherwise the VB's
'  reference count would be incorrect)
Public Sub DestroyRef(nobj As Long)
    Dim N As Long
    CopyMemory nobj, VarPtr(N), 4
End Sub

'The window procedure for the dummy windows that clsInTray creates
Public Function InTrayWndProc(ByVal hwnd As Long, ByVal uMsg As Long, _
       ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Dim lpObj As Long
    Dim obj As clsInTray
    
    If Not m_bMessageInited Then
        InitMessage
    End If
    
    Select Case uMsg
        'Pass WM_USER_SYSTRAY to the clsInTray object
        Case WM_USER_SYSTRAY
            lpObj = GetWindowLong(hwnd, GWL_USERDATA)
            Set obj = DeRef(lpObj)
            obj.ProcessMessage wParam, lParam
            DestroyRef VarPtr(obj)
        'If the TaskBar restarts, let clsInTray know about it
        Case m_uTaskbarRestart
            lpObj = GetWindowLong(hwnd, GWL_USERDATA)
            Set obj = DeRef(lpObj)
            obj.RecreateIcon
            DestroyRef VarPtr(obj)
    End Select
    
    InTrayWndProc = DefWindowProc(hwnd, uMsg, wParam, lParam)
    
End Function

'Register the windows message TaskbarCreated so we can watch for it
Private Function InitMessage()
    
    m_bMessageInited = True
    
    m_uTaskbarRestart = RegisterWindowMessage("TaskbarCreated")
    
End Function

'End of Module: modInTray
'----------------------------------------------------------------------

