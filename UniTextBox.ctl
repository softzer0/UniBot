VERSION 5.00
Begin VB.UserControl UniTextBox 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2655
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   20
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   177
End
Attribute VB_Name = "UniTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------------------------- APIs -------------------------------
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowLongW Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CreateFont Lib "GDI32" Alias "CreateFontA" (ByVal h As Long, ByVal w As Long, ByVal E As Long, ByVal O As Long, ByVal w As Long, ByVal i As Long, ByVal U As Long, ByVal s As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As String) As Long
Private Declare Function SendMessageW Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowTextLength& Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long)
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long

' for subclassing
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long

'---------------------------------- Private Constants -------------------------------
' for subclassing
Private Const ALL_MESSAGES          As Long = -1                                       'All messages added or deleted
Private Const MEM_COMMIT = &H1000&, PAGE_RWX = &H40&, MEM_RELEASE = &H8000&
Private Const GWL_WNDPROC           As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04              As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05              As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08              As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09              As Long = 137

'Mouse & Key Event
Private Const WM_SETFOCUS            As Long = &H7
Private Const WM_LBUTTONUP          As Long = &H202
Private Const WM_LBUTTONDOWN        As Long = &H201
Private Const WM_LBUTTONDBLCLK      As Long = &H203
Private Const WM_RBUTTONDBLCLK      As Long = &H206
Private Const WM_RBUTTONDOWN        As Long = &H204
Private Const WM_RBUTTONUP          As Long = &H205
Private Const WM_MOUSEMOVE          As Long = &H200
Private Const WM_MOUSELEAVE         As Long = &H2A3
Private Const WM_CHAR               As Long = &H102
Private Const WM_VSCROLL            As Long = &H115

'Style
'Private Const GWL_STYLE = (-16)
'Private Const GWL_EXSTYLE As Long = -20
Private Const WS_CHILD As Long = &H40000000
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_VSCROLL As Long = &H200000
Private Const WS_HSCROLL As Long = &H100000
Private Const WS_BORDER As Long = &H800000
Private Const WS_TABSTOP As Long = &H10000
Private Const ES_MULTILINE As Long = &H4&
Private Const ES_WANTRETURN As Long = &H1000&
Private Const ES_AUTOVSCROLL As Long = &H40&
Private Const ES_AUTOHSCROLL As Long = &H80&
Private Const ES_READONLY As Long = &H800&
Private Const ES_CENTER& = &H1&
Private Const ES_LEFT& = &H0&
Private Const ES_RIGHT& = &H2&
Private Const ES_PASSWORD As Long = &H20&
Private Const WS_MAXIMIZEBOX = &H10000
Private Const ES_NOHIDESEL As Long = &H100&
Private Const ES_NUMBER As Long = &H2000&
Private Const ES_LOWERCASE As Long = &H10&
Private Const ES_UPPERCASE As Long = &H8&
'-------------------------------------------
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_CLIPCHILDREN = &H20000004
Private Const WS_MAXIMIZE = &H1000000
' ExWindowStyles
'Private Const WS_EX_ACCEPTFILES = &H10&
Private Const WS_EX_CLIENTEDGE As Long = &H200&

'SendMessage
Private Const VK_RETURN As Long = &HD
Private Const VK_SPACE As Long = &H20
Private Const WM_KEYUP = &H101
Private Const WM_KEYDOWN = &H100
Private Const WM_SETTEXT As Long = &HC
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_GETFONT As Long = &H31
Private Const WM_SETFONT As Long = &H30
Private Const EM_SETPASSWORDCHAR = &HCC
Private Const EM_SETREADONLY As Long = &HCF
Private Const EM_GETSEL As Long = &HB0
'Private Const EM_GETLINE As Long = &HC4
Private Const EM_LIMITTEXT As Long = &HC5
Private Const EM_SETSEL As Long = &HB1
Private Const EM_REPLACESEL As Long = &HC2
Private Const WM_USER As Long = &H400
Private Const EM_GETTEXTRANGE As Long = (WM_USER + 75)
Private Const SB_BOTTOM As Long = 7

'---------------------------------- Enums -------------------------------
' for subclassing
Private Enum eMsgWhen
    MSG_AFTER = 1                                                                         'Message calls back after the original (previous) WndProc
    MSG_BEFORE = 2                                                                        'Message calls back before the original (previous) WndProc
    MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                        'Message calls back before and after the original (previous) WndProc
End Enum

Public Enum SCROLL_BAR
    None = 0
    Horizontal = 1
    Vertical = 2
    both = 3
End Enum

Public Enum TEXT_CONVERT
    [None] = 0
    [lowercase] = 1
    [UPPERCASE] = 2
End Enum

Public Enum APPEAR
    [Flat] = 0
    [3D] = 1
End Enum

Public Enum BORDER_STYLE
    [None] = 0
    [Fixed Single] = 1
End Enum

Public Enum TEXT_ALIGN
    [Left Justify] = 0
    [Right Justify] = 1
    [Center] = 2
End Enum

Private Type tSubData
    hwnd          As Long
    nAddrSub      As Long
    nAddrOrig     As Long
    nMsgCntA      As Long
    nMsgCntB      As Long
    aMsgTblA()    As Long
    aMsgTblB()    As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private sc_aSubData()               As tSubData
Private hEdit                       As Long
Private m_Text                      As String
Private m_MultiLine                 As Boolean
Private m_ReadOnly                  As Boolean
Private m_ScrollBar                 As SCROLL_BAR
Private m_BorderStyle               As BORDER_STYLE
Private m_PasswordChar              As String
Private m_TextAlign                 As TEXT_ALIGN
Private m_SelectText                As Boolean
Private m_Enabled                   As Boolean
Private m_appear                    As APPEAR
Private m_hideSel                   As Boolean
Private m_TextConvert               As TEXT_CONVERT
Private m_numberOnly                As Boolean
Private m_maxLen                    As Long
Private m_SelLength                 As Long
Private m_SelStart                  As Long
Private m_SelText                   As String

'[Events]
Public Event Change()
Public Event Click()
Public Event DbClick()
'Public Event SelChange()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
'
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'
Public Event MouseLeave()
'
'Sublassing Control
'
Public Sub zSubclass_Text_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)

'****************Tim vitri chuot
Dim ToadoX As Single, ToadoY As Single
ToadoX = WordLo(lParam) * 15
ToadoY = WordHi(lParam) * 15
'*********************
  
    If Ambient.UserMode = False Then Subclass_StopAll:  Exit Sub
    
    Select Case lng_hWnd
        Case UserControl.hwnd
            Select Case uMsg
                Case WM_SETFOCUS
                    Me.Focus

            End Select
            
        Case hEdit
            Select Case uMsg

                Case WM_KEYDOWN
                    RaiseEvent KeyDown(wParam And &H7FFF&, pvShiftState())
                    
                Case WM_CHAR
                    If ((wParam And &H7FFF&) = vbKeyTab) Then Putfocus UserControl.Parent.hwnd

                    RaiseEvent KeyPress(wParam And &H7FFF&)

                Case WM_KEYUP
                    RaiseEvent Change
                    RaiseEvent KeyUp(wParam And &H7FFF&, pvShiftState())
                    
                Case WM_RBUTTONUP
                    RaiseEvent MouseUp(wParam And &H7FFF&, pvShiftState(), ToadoX, ToadoY)
'                    RaiseEvent Click
                    
                Case WM_RBUTTONDOWN
                    RaiseEvent MouseDown(wParam And &H7FFF&, pvShiftState(), ToadoX, ToadoY)
                    
                Case WM_LBUTTONUP
                    RaiseEvent MouseUp(wParam And &H7FFF&, pvShiftState(), ToadoX, ToadoY)
                    RaiseEvent Click
                    
                Case WM_LBUTTONDOWN
                    RaiseEvent MouseDown(wParam And &H7FFF&, pvShiftState(), ToadoX, ToadoY)
                    
                Case WM_MOUSEMOVE
                    RaiseEvent MouseMove(wParam And &H7FFF&, pvShiftState(), ToadoX, ToadoY)
                    
                Case WM_LBUTTONDBLCLK
                    RaiseEvent DbClick
                    
                Case WM_RBUTTONDBLCLK
                    RaiseEvent DbClick
                    
                Case WM_MOUSELEAVE
                    RaiseEvent MouseLeave
            End Select
    End Select
End Sub

Private Sub InitializeSubClassing()
    On Error GoTo handle
'-- Subclass UserControl (parent)
        With UserControl
            Call Subclass_Start(.hwnd)
            Call Subclass_AddMsg(.hwnd, WM_SETFOCUS, MSG_AFTER)
        End With

'-- Subclass UniXPTextBox (child)
        Call Subclass_Start(hEdit)
        Call Subclass_AddMsg(hEdit, WM_CHAR, MSG_BEFORE)
        Call Subclass_AddMsg(hEdit, WM_KEYUP, MSG_AFTER)
        Call Subclass_AddMsg(hEdit, WM_KEYDOWN, MSG_AFTER)
        
        Call Subclass_AddMsg(hEdit, WM_LBUTTONDOWN, MSG_AFTER)
        Call Subclass_AddMsg(hEdit, WM_LBUTTONUP, MSG_AFTER)
        Call Subclass_AddMsg(hEdit, WM_RBUTTONDOWN, MSG_AFTER)
        Call Subclass_AddMsg(hEdit, WM_RBUTTONUP, MSG_AFTER)
        
        Call Subclass_AddMsg(hEdit, WM_LBUTTONDBLCLK, MSG_AFTER)
        Call Subclass_AddMsg(hEdit, WM_RBUTTONDBLCLK, MSG_AFTER)
        
        Call Subclass_AddMsg(hEdit, WM_MOUSEMOVE, MSG_AFTER)
        Call Subclass_AddMsg(hEdit, WM_MOUSELEAVE, MSG_AFTER)
handle:
End Sub

'---------------------------------------------------
'Properties
'---------------------------------------------------

Public Property Get TextConvert() As TEXT_CONVERT
    TextConvert = m_TextConvert
End Property

Public Property Let TextConvert(new_value As TEXT_CONVERT)
    m_TextConvert = new_value
    PropertyChanged "TextConvert"
    DoRefresh
End Property

Public Property Get SelLength() As Long
Attribute SelLength.VB_MemberFlags = "400"
    Dim pos As Long
    pos = SendMessage(hEdit, EM_GETSEL, 0&, 0&)
    SelLength = WordHi(pos) - WordLo(pos)
End Property

Public Property Let SelLength(ByVal new_value As Long)
    m_SelLength = new_value
    Call SendMessage(hEdit, EM_SETSEL, m_SelStart, ByVal m_SelLength + m_SelStart)
    PropertyChanged "SelLength"
End Property

Public Property Get SelStart() As Long
Attribute SelStart.VB_MemberFlags = "400"
    Dim pos As Long
    pos = SendMessage(hEdit, EM_GETSEL, 0&, 0&)
    SelStart = WordLo(pos) + 1
End Property

Public Property Let SelStart(ByVal new_value As Long)
    m_SelStart = new_value
    Call SendMessage(hEdit, EM_SETSEL, m_SelStart, ByVal m_SelLength + m_SelStart)
    PropertyChanged "SelStart"
End Property

Public Property Get SelText() As String
Attribute SelText.VB_MemberFlags = "400"
    m_Text = TextBox_GetText(hEdit)
    Dim pos As Long
    pos = SendMessage(hEdit, EM_GETSEL, 0&, 0&)
    SelText = Mid(m_Text, WordLo(pos) + 1, WordHi(pos) - WordLo(pos))
End Property

Public Property Let SelText(ByVal new_value As String)
    Subclass_StopAll
    Call SendMessageW(hEdit, EM_REPLACESEL, 0&, StrPtr(new_value))
    InitializeSubClassing
    PropertyChanged "SelText"
End Property

Public Property Get MaxLength() As Long
    MaxLength = m_maxLen
End Property

Public Property Let MaxLength(new_value As Long)
    m_maxLen = new_value
    Call SendMessage(hEdit, EM_LIMITTEXT, m_maxLen, 0&)
    PropertyChanged "MaxLength"
End Property

Public Property Get NumberOnly() As Boolean
    NumberOnly = m_numberOnly
End Property

Public Property Let NumberOnly(new_value As Boolean)
    m_numberOnly = new_value
    PropertyChanged "NumberOnly"
    DoRefresh
End Property

Public Property Get HideSelection() As Boolean
    HideSelection = m_hideSel
End Property

Public Property Let HideSelection(new_value As Boolean)
    m_hideSel = new_value
    PropertyChanged "HideSelection"
    DoRefresh
End Property
    
Public Property Get Alignment() As TEXT_ALIGN
    Alignment = m_TextAlign
End Property

Public Property Let Alignment(New_Align As TEXT_ALIGN)
    m_TextAlign = New_Align
    PropertyChanged "Alignment"
    DoRefresh
End Property

Public Property Get Locked() As Boolean
     Locked = m_ReadOnly
End Property

Public Property Let Locked(new_value As Boolean)
    m_ReadOnly = new_value
    PropertyChanged "Locked"
    'DoRefresh
    SendMessage hEdit, EM_SETREADONLY, m_ReadOnly, 0&
End Property

Public Property Get Multiline() As Boolean
     Multiline = m_MultiLine
End Property

Public Property Let Multiline(new_value As Boolean)
    m_MultiLine = new_value
    PropertyChanged "MultiLine"
    DoRefresh
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(new_value As Boolean)
    m_Enabled = new_value
    UserControl.Enabled = new_value
    PropertyChanged "Enabled"
    DoRefresh
End Property

Public Property Get Scrollbar() As SCROLL_BAR
    Scrollbar = m_ScrollBar
End Property

Public Property Let Scrollbar(New_ScrollBar As SCROLL_BAR)
    m_ScrollBar = New_ScrollBar
    PropertyChanged "Scrollbar"
    DoRefresh
End Property

Public Property Get Text() As String
     m_Text = TextBox_GetText(hEdit)
     Text = m_Text
End Property ' Get Caption

Public Property Let Text(ByVal new_value As String)
     m_Text = TextBox_SetText(hEdit, new_value)
     PropertyChanged "Text"
     RaiseEvent Change
End Property ' Let Caption

Public Property Get Font() As StdFont
     Set Font = UserControl.Font
End Property ' Get Font

Public Property Set Font(ByVal new_Font As StdFont)
    Set UserControl.Font = new_Font
    PropertyChanged "Font"
    SetFont
End Property ' Let Font

Public Property Get Appearance() As APPEAR
    Appearance = m_appear
End Property

Public Property Let Appearance(new_value As APPEAR)
    m_appear = new_value
    UserControl.Appearance = m_appear
    PropertyChanged "Appearance"
    DoRefresh
End Property

Public Property Get BorderStyle() As BORDER_STYLE
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(new_BorderStyle As BORDER_STYLE)
    m_BorderStyle = new_BorderStyle
    PropertyChanged "BorderStyle"
    DoRefresh
End Property

Public Property Get BackColor() As OLE_COLOR
   BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal new_Color As OLE_COLOR)
    UserControl.BackColor = new_Color
    PropertyChanged "BackColor"
    'SendMessageLong hEdit, EM_SETBKGNDCOLOR, 0, TranslateColor(new_Color)
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(New_ForeColor As OLE_COLOR)
    UserControl.ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    SetFont
End Property

Public Property Get PasswordChar() As String
    PasswordChar = m_PasswordChar
End Property

Public Property Let PasswordChar(new_PasswordChar As String)
    m_PasswordChar = Left(new_PasswordChar, 1)
    If m_PasswordChar <> "" And Not m_MultiLine Then
        SendMessage hEdit, EM_SETPASSWORDCHAR, ByVal AscW(m_PasswordChar), 0
    Else
        SendMessage hEdit, EM_SETPASSWORDCHAR, ByVal 0&, 0
    End If
    UserControl.Refresh
    PropertyChanged "PasswordChar"
End Property
'----------------- End Properties ------------------

'---------------------------------------------------
'Functions
'---------------------------------------------------

Private Function CreateTextBox(hParent As Long, strCaption As String, x As Long, y As Long, Width As Long, Height As Long, _
                                Optional Scroll As SCROLL_BAR = 0, _
                                Optional Align As TEXT_ALIGN = [Left Justify], _
                                Optional Convert As TEXT_CONVERT = 0, _
                                Optional Multiline As Boolean = False, _
                                Optional HideSel As Boolean = False, _
                                Optional NumberOnly As Boolean = False, _
                                Optional password As Boolean = False, _
                                Optional readOnly As Boolean = False, _
                                Optional lAdditionalStyles As Long = 0) As Long
    
    If hParent = 0 Then Exit Function
    
    Dim lExStyle As Long, lStyle As Long
    
    '-- Define window style

    lStyle = WS_CHILD Or WS_TABSTOP Or WS_VISIBLE Or WS_CLIPSIBLINGS 'Or ES_WANTRETURN  'Or ES_AUTOVSCROLL Or ES_AUTOHSCROLL

    If m_BorderStyle = [Fixed Single] Then
        If m_appear = [3D] Then lExStyle = WS_EX_CLIENTEDGE Else lStyle = lStyle Or WS_BORDER
    End If

    If Multiline Then
        Select Case Scroll
'        Case [None]
'            lStyle = lStyle
        Case [Horizontal]
            lStyle = lStyle Or WS_HSCROLL
        Case [Vertical]
            lStyle = lStyle Or WS_VSCROLL
        Case [both]
            lStyle = lStyle Or WS_VSCROLL Or WS_HSCROLL
        End Select
        lStyle = lStyle Or ES_MULTILINE Or ES_AUTOVSCROLL
    Else
        lStyle = lStyle Or ES_AUTOHSCROLL
    End If
    
    If HideSel Then lStyle = lStyle Or ES_NOHIDESEL
    
    If readOnly Then lStyle = lStyle Or ES_READONLY
    
    If NumberOnly Then lStyle = lStyle Or ES_NUMBER
    
    If password And Not Multiline And m_PasswordChar = "*" Then
        lStyle = lStyle Or ES_PASSWORD
    End If
    
    Select Case Align
        Case [Left Justify]
            lStyle = lStyle Or ES_LEFT
        Case [Right Justify]
            lStyle = lStyle Or ES_RIGHT
        Case Center
            lStyle = lStyle Or ES_CENTER
    End Select

    Select Case Convert
'        Case [None]
'            lStyle = lStyle
        Case [lowercase]
            lStyle = lStyle Or ES_LOWERCASE
        Case [UPPERCASE]
            lStyle = lStyle Or ES_UPPERCASE
    End Select
    
    If lAdditionalStyles > 0 Then lStyle = lStyle Or lAdditionalStyles

    Dim hTemp As Long
    hTemp = CreateWindowExW(lExStyle, StrPtr("eDit"), StrPtr(strCaption), lStyle, x, y, Width, Height, hParent, vbNull, App.hInstance, ByVal 0&)
    If hTemp <= 0 Then Exit Function
    CreateTextBox = hTemp
    
    If password And Not Multiline And m_PasswordChar <> "*" Then
        SendMessage hTemp, EM_SETPASSWORDCHAR, ByVal AscW(m_PasswordChar), 0
    ElseIf Not password Then
        SendMessage hTemp, EM_SETPASSWORDCHAR, ByVal 0&, 0
    End If

End Function

Private Sub DoRefresh()
    
    If hEdit <> 0 Then m_Text = TextBox_GetText(hEdit)
    
    UserControl_Terminate
    
    hEdit = CreateTextBox(UserControl.hwnd, m_Text, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, m_ScrollBar, m_TextAlign, m_TextConvert, m_MultiLine, m_hideSel, m_numberOnly, Len(m_PasswordChar), m_ReadOnly)
    InitializeSubClassing
    TextBox_SetText hEdit, m_Text
    SetFont
    
    SetWindowPos hEdit, 0, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 0
    
    'InitializeSubClassing
End Sub

Public Sub ScrollToBottom()
   Subclass_StopAll
   SendMessage hEdit, WM_VSCROLL, SB_BOTTOM, 0&
   InitializeSubClassing
End Sub

Private Function TextBox_SetText(hwnd As Long, sText As String) As Long
    Subclass_StopAll
    TextBox_SetText = SendMessageW(hwnd, WM_SETTEXT, 0&, ByVal StrPtr(sText))
    InitializeSubClassing
End Function

Private Function TextBox_GetText(hwnd As Long) As String
    Dim sText As String
    Dim lLength As Long
    Subclass_StopAll
    ' Get the text's length.
    lLength = SendMessage(hwnd, WM_GETTEXTLENGTH, 0, ByVal 0&)
    sText = Space$(lLength + 1)

    ' Get the text.
    Call SendMessageW(hwnd, WM_GETTEXT, lLength + 1, ByVal StrPtr(sText))
    TextBox_GetText = Left$(sText, lLength)
    InitializeSubClassing
End Function

Private Sub SetFont()
Dim hFont As Long
    hFont = SendMessage(UserControl.hwnd, WM_GETFONT, 0, 0)
    SendMessage hEdit, WM_SETFONT, hFont, 0
End Sub

Public Sub Focus()
    If hEdit <> 0 Then Putfocus hEdit
End Sub

Private Sub UserControl_EnterFocus()
  'Call SetFocus(hEdit)
End Sub

'-----------------------------------------------------
'UserControl's Events
'-----------------------------------------------------

Private Sub UserControl_InitProperties()
   
    Set UserControl.Font = Ambient.Font
    m_Text = Extender.Name
    m_MultiLine = False
    m_ReadOnly = False
    m_ScrollBar = 0
    m_BorderStyle = [Fixed Single]
    m_PasswordChar = ""
    m_TextAlign = [Left Justify]
    m_Enabled = True
    m_appear = [3D]
    m_hideSel = False
    m_TextConvert = 0
    m_numberOnly = False
    m_maxLen = 0
    
    DoRefresh

End Sub

Private Sub UserControl_Resize()
    SetWindowPos hEdit, 0, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, 0
End Sub

Private Sub UserControl_Terminate()
On Error GoTo Catch
    If hEdit <> 0 Then
        DestroyWindow hEdit
        Call Subclass_StopAll
    End If
Catch:
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    With PropBag
        Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
        UserControl.ForeColor = .ReadProperty("ForeColor", Ambient.ForeColor)
        UserControl.BackColor = .ReadProperty("BackColor", Ambient.BackColor)
        m_Text = .ReadProperty("Text", Ambient.DisplayName)
        m_PasswordChar = .ReadProperty("PasswordChar", "")
        m_MultiLine = .ReadProperty("MultiLine", False)
        m_appear = .ReadProperty("Appearance", [3D])
        m_ReadOnly = .ReadProperty("Locked", False)
        m_Enabled = .ReadProperty("Enabled", True)
        UserControl.Enabled = m_Enabled
        m_BorderStyle = .ReadProperty("BorderStyle", [Fixed Single])
        m_ScrollBar = .ReadProperty("Scrollbar", 0)
        m_TextAlign = .ReadProperty("Alignment", [Left Justify])
        m_hideSel = .ReadProperty("HideSelection", False)
        m_TextConvert = .ReadProperty("TextConvert", 0)
        m_numberOnly = .ReadProperty("NumberOnly", False)
        m_maxLen = .ReadProperty("MaxLength", 0)
    End With

    DoRefresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        Call .WriteProperty("Font", UserControl.Font, Ambient.Font)
        Call .WriteProperty("ForeColor", UserControl.ForeColor, Ambient.ForeColor)
        Call .WriteProperty("BackColor", UserControl.BackColor, Ambient.BackColor)
        Call .WriteProperty("Text", m_Text, Extender.Name)
        Call .WriteProperty("PasswordChar", m_PasswordChar, "")
        Call .WriteProperty("MultiLine", m_MultiLine, False)
        Call .WriteProperty("Appearance", m_appear, [3D])
        Call .WriteProperty("Locked", m_ReadOnly, False)
        Call .WriteProperty("Enabled", m_Enabled, True)
        Call .WriteProperty("BorderStyle", m_BorderStyle, [Fixed Single])
        Call .WriteProperty("Scrollbar", m_ScrollBar, 0)
        Call .WriteProperty("Alignment", m_TextAlign, [Left Justify])
        Call .WriteProperty("HideSelection", m_hideSel, False)
        Call .WriteProperty("TextConvert", m_TextConvert, 0)
        Call .WriteProperty("NumberOnly", m_numberOnly, False)
        Call .WriteProperty("MaxLength", m_maxLen, 0)
    End With
End Sub
'------------- End UserControl ---------------

'---------------------------------------------
'------------- SubClassing -------------------
'---------------------------------------------
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    With sc_aSubData(zIdx(lng_hWnd))
        If When And eMsgWhen.MSG_BEFORE Then
            Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If When And eMsgWhen.MSG_AFTER Then
            Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    With sc_aSubData(zIdx(lng_hWnd))
        If When And eMsgWhen.MSG_BEFORE Then
            Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If When And eMsgWhen.MSG_AFTER Then
            Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
    Const CODE_LEN              As Long = 200
    Const FUNC_CWP              As String = "CallWindowProcA"
    Const FUNC_EBM              As String = "EbMode"
    Const FUNC_SWL              As String = "SetWindowLongA"
    Const MOD_USER              As String = "user32"
    Const MOD_VBA5              As String = "vba5"
    Const MOD_VBA6              As String = "vba6"
    Const PATCH_01              As Long = 18
    Const PATCH_02              As Long = 68
    Const PATCH_03              As Long = 78
    Const PATCH_06              As Long = 116
    Const PATCH_07              As Long = 121
    Const PATCH_0A              As Long = 186
    Static aBuf(1 To CODE_LEN)  As Byte
    Static pCWP                 As Long
    Static pEbMode              As Long
    Static pSWL                 As Long
    Dim i                       As Long
    Dim j                       As Long
    Dim nSubIdx                 As Long
    Dim sHex                    As String
  
    If aBuf(1) = 0 Then
  
        sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
            "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
            "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
            "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
        
        i = 1
        Do While j < CODE_LEN
            j = j + 1
            aBuf(j) = Val("&H" & Mid$(sHex, i, 2))
            i = i + 2
        Loop
    
        On Error Resume Next
        Debug.Assert 1 / 0
        If err Then
          err.Clear
          pCWP = zAddrFunc(MOD_VBA6, FUNC_EBM)
        End If
        On Error GoTo 0
            
        pCWP = zAddrFunc(MOD_USER, FUNC_CWP)
        pSWL = zAddrFunc(MOD_USER, FUNC_SWL)
        ReDim sc_aSubData(0 To 0) As tSubData
    Else
        nSubIdx = zIdx(lng_hWnd, True)
        If nSubIdx = -1 Then
            nSubIdx = UBound(sc_aSubData()) + 1
            ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData
        End If
        Subclass_Start = nSubIdx
    End If

    With sc_aSubData(nSubIdx)
        .hwnd = lng_hWnd
        .nAddrSub = VirtualAlloc(0&, CODE_LEN, MEM_COMMIT, PAGE_RWX)
        .nAddrOrig = SetWindowLong(.hwnd, GWL_WNDPROC, .nAddrSub)
        Call RtlMoveMemory(ByVal .nAddrSub, aBuf(1), CODE_LEN)
        Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)
        Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)
        Call zPatchRel(.nAddrSub, PATCH_03, pSWL)
        Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)
        Call zPatchRel(.nAddrSub, PATCH_07, pCWP)
        Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))
    End With
End Function

Private Sub Subclass_StopAll()
On Error GoTo err
    Dim i As Long
    
    i = UBound(sc_aSubData())                                                             'Get the upper bound of the subclass data array
    Do While i >= 0                                                                       'Iterate through each element
        With sc_aSubData(i)
            If .hwnd <> 0 Then                                                                'If not previously Subclass_Stop'd
                Call Subclass_Stop(.hwnd)                                                       'Subclass_Stop
            End If
        End With
        i = i - 1                                                                           'Next element
    Loop
err:
End Sub

Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
  With sc_aSubData(zIdx(lng_hWnd))
    Call SetWindowLong(.hwnd, GWL_WNDPROC, .nAddrOrig)                                 'Restore the original WndProc
    Call zPatchVal(.nAddrSub, PATCH_05, 0)                                              'Patch the Table B entry count to ensure no further 'before' callbacks
    Call zPatchVal(.nAddrSub, PATCH_09, 0)                                              'Patch the Table A entry count to ensure no further 'after' callbacks
    Call VirtualFree(.nAddrSub, 0&, MEM_RELEASE)                                                          'Release the machine code memory
    .hwnd = 0                                                                           'Mark the sc_aSubData element as available for re-use
    .nMsgCntB = 0                                                                       'Clear the before table
    .nMsgCntA = 0                                                                       'Clear the after table
    Erase .aMsgTblB                                                                     'Erase the before table
    Erase .aMsgTblA                                                                     'Erase the after table
  End With
End Sub

Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
  Dim nEntry  As Long                                                                   'Message table entry index
  Dim nOff1   As Long                                                                   'Machine code buffer offset 1
  Dim nOff2   As Long                                                                   'Machine code buffer offset 2
  
  If uMsg = ALL_MESSAGES Then                                                           'If all messages
    nMsgCnt = ALL_MESSAGES                                                              'Indicates that all messages will callback
  Else                                                                                  'Else a specific message number
    Do While nEntry < nMsgCnt                                                           'For each existing entry. NB will skip if nMsgCnt = 0
      nEntry = nEntry + 1
      
      If aMsgTbl(nEntry) = 0 Then                                                       'This msg table slot is a deleted entry
        aMsgTbl(nEntry) = uMsg                                                          'Re-use this entry
        Exit Sub                                                                        'Bail
      ElseIf aMsgTbl(nEntry) = uMsg Then                                                'The msg is already in the table!
        Exit Sub                                                                        'Bail
      End If
    Loop                                                                                'Next entry

    nMsgCnt = nMsgCnt + 1                                                               'New slot required, bump the table entry count
    ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                        'Bump the size of the table.
    aMsgTbl(nMsgCnt) = uMsg                                                             'Store the message number in the table
  End If

  If When = eMsgWhen.MSG_BEFORE Then                                                    'If before
    nOff1 = PATCH_04                                                                    'Offset to the Before table
    nOff2 = PATCH_05                                                                    'Offset to the Before table entry count
  Else                                                                                  'Else after
    nOff1 = PATCH_08                                                                    'Offset to the After table
    nOff2 = PATCH_09                                                                    'Offset to the After table entry count
  End If

  If uMsg <> ALL_MESSAGES Then
    Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                                    'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
  End If
  Call zPatchVal(nAddr, nOff2, nMsgCnt)                                                 'Patch the appropriate table entry count
End Sub

Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
End Function

Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    Dim nEntry As Long
    
    If uMsg = ALL_MESSAGES Then
        nMsgCnt = 0
        If When = eMsgWhen.MSG_BEFORE Then
            nEntry = PATCH_05
        Else
            nEntry = PATCH_09
        End If
        Call zPatchVal(nAddr, nEntry, 0)
    Else
        Do While nEntry < nMsgCnt
            nEntry = nEntry + 1
            If aMsgTbl(nEntry) = uMsg Then
                aMsgTbl(nEntry) = 0
                Exit Do
            End If
        Loop
    End If
End Sub

Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long

    zIdx = UBound(sc_aSubData)
    Do While zIdx >= 0
        With sc_aSubData(zIdx)
        If .hwnd = lng_hWnd Then
            If Not bAdd Then
            Exit Function
            End If
        ElseIf .hwnd = 0 Then
            If bAdd Then
            Exit Function
            End If
        End If
        End With
        zIdx = zIdx - 1
    Loop
End Function

Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
    zSetTrue = True
    bValue = True
End Function

Private Function WordHi(lngValue As Long) As Long
    If (lngValue And &H80000000) = &H80000000 Then
        WordHi = ((lngValue And &H7FFF0000) \ &H10000) Or &H8000&
    Else
        WordHi = (lngValue And &HFFFF0000) \ &H10000
    End If
End Function

Private Function WordLo(lngValue As Long) As Long
    WordLo = (lngValue And &HFFFF&)
End Function
'------------- End Subclassing ---------------

Private Function pvShiftState() As Integer

  Dim lS As Integer
    If (GetAsyncKeyState(vbKeyShift) < 0) Then
        lS = lS Or vbShiftMask
    End If
    If (GetAsyncKeyState(vbKeyMenu) < 0) Then
        lS = lS Or vbAltMask
    End If
    If (GetAsyncKeyState(vbKeyControl) < 0) Then
        lS = lS Or vbCtrlMask
    End If
    pvShiftState = lS
End Function
