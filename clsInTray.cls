VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------
'Start of Class Module: clsInTray

Option Explicit

Private m_hWndHidden As Long 'Dummy window used for call backs
Private m_nid As NOTIFYICONDATA
Private m_bVisible As Boolean 'Is the icon visible yet?
Private m_sClass As String

Public Event MouseMove()
Public Event MouseUp(Button As Integer)
Public Event MouseDown(Button As Integer)
Public Event DblClick(Button As Integer)

'The handle of the hIcon to use
Public Property Get hIcon() As Long
    
    hIcon = m_nid.hIcon
    
End Property

Public Property Let hIcon(newhIcon As Long)
    
    m_nid.hIcon = newhIcon
    
End Property

'Tool tip for the system tray icon
Public Property Get Tip() As String
    
    Tip = m_nid.szTip
    
End Property

Public Property Let Tip(szTip As String)
    
    m_nid.szTip = szTip & vbNullChar
    
End Property

'Adds the icon to the system tray.
' Set hIcon and Tip before calling this function
Public Sub AddIcon()
    
    Debug.Assert m_bVisible = False
    
    Shell_NotifyIcon NIM_ADD, m_nid
    
    m_bVisible = True
    
End Sub

'Updates the system tray icon to reflect any changes to
' hIcon or Tip
Public Sub ModifyIcon()
    
    Debug.Assert m_bVisible
    
    Shell_NotifyIcon NIM_MODIFY, m_nid
    
End Sub

'Removes the system tray icon
Public Sub RemoveIcon()
    
    Debug.Assert m_bVisible
    
    Shell_NotifyIcon NIM_DELETE, m_nid
    
    m_bVisible = False
    
End Sub

Private Sub Class_Initialize()
    
    m_sClass = "VB InTray Class"
    
    'Register a class for, and create the dummy window
    Dim wc As WNDCLASS
    wc.style = 0
    wc.lpfnwndproc = Pass(AddressOf InTrayWndProc)
    wc.hInstance = App.hInstance
    wc.lpszClassName = m_sClass
    RegisterClass wc
    m_hWndHidden = CreateWindowEx(0, m_sClass, "VB InTray Hidde" & _
                 "n Window", 0, 0, 0, 100, 100, 0, 0, 0, ByVal 0)
    
    Dim lpMe As Long
    lpMe = CreateRef(Me)
    SetWindowLong m_hWndHidden, GWL_USERDATA, lpMe
    
    'Initialize the Notify Icon Data structure
    With m_nid
        .cbSize = Len(m_nid)
        .hIcon = 0
        .hwnd = m_hWndHidden
        .szTip = vbNullString & vbNullString
        .uCallbackMessage = WM_USER_SYSTRAY
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uID = 0
    End With
    
End Sub

'Function should only be called from modInTray to process windows
' messages generated from the System Tray
Public Sub ProcessMessage(wParam As Long, lParam As Long)
    
    Select Case lParam
        Case WM_MOUSEMOVE
            RaiseEvent MouseMove
        Case WM_LBUTTONDOWN
            RaiseEvent MouseDown(vbLeftButton)
        Case WM_LBUTTONUP
            RaiseEvent MouseUp(vbLeftButton)
        Case WM_LBUTTONDBLCLK
            RaiseEvent DblClick(vbLeftButton)
        Case WM_RBUTTONDOWN
            RaiseEvent MouseDown(vbRightButton)
        Case WM_RBUTTONUP
            RaiseEvent MouseUp(vbRightButton)
        Case WM_RBUTTONDBLCLK
            RaiseEvent DblClick(vbRightButton)
        Case WM_MBUTTONDOWN
            RaiseEvent MouseDown(vbMiddleButton)
        Case WM_MBUTTONUP
            RaiseEvent MouseUp(vbMiddleButton)
        Case WM_MBUTTONDBLCLK
            RaiseEvent DblClick(vbMiddleButton)
    End Select
    
End Sub

Private Sub Class_Terminate()
    
    'Close the dummy window
    SendMessage m_hWndHidden, WM_CLOSE, 0, ByVal 0&
    
    'Remove the icon if it exists
    If m_bVisible Then
        RemoveIcon
    End If
    
    UnregisterClass m_sClass, App.hInstance

End Sub

'This function should only be called from modInTray when the
' TaskBar is recreated (ie, in response to TaskbarCreated)
Public Sub RecreateIcon()
    
    If m_bVisible Then
        m_bVisible = False
        AddIcon
    End If
    
End Sub

'End of Class Module: clsInTray
'----------------------------------------------------------------------
