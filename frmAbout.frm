VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   1508
   ClientLeft      =   2340
   ClientTop       =   1937
   ClientWidth     =   2938
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1045.68
   ScaleMode       =   0  'User
   ScaleWidth      =   2760.812
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   676
      Left            =   120
      Picture         =   "frmAbout.frx":000C
      ScaleHeight     =   434.95
      ScaleMode       =   0  'User
      ScaleWidth      =   434.95
      TabIndex        =   0
      Top             =   120
      Width           =   676
   End
   Begin VB.Label lblSite 
      BackStyle       =   0  'Transparent
      Caption         =   "unibot.boards.net"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.51
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   1502
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblSite 
      BackStyle       =   0  'Transparent
      Caption         =   "www.mikisoft.me"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.51
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   260
      Index           =   0
      Left            =   1508
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   962
      Width           =   1313
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.51
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   1845
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.51
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1080
      TabIndex        =   3
      Top             =   360
      Width           =   1605
   End
   Begin VB.Label Label1 
      Caption         =   "Visit:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.51
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      ToolTipText     =   "And check for updates at:"
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label lblDescription 
      Caption         =   "By MikiSoft"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.51
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   1005
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type
Private Type GUID
    data1 As Long
    data2 As Integer
    data3 As Integer
    data4(0 To 7) As Byte
End Type
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" _
    (lpPictDesc As PictDesc, riid As GUID, _
    ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
Private Const IDC_HAND = 32649&
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long

Dim lol As Boolean

Public Function IconToPicture(ByVal hIcon As Long) As IPicture
    
    If hIcon = 0 Then Exit Function
        
    Dim oNewPic As Picture
    Dim tPicConv As PictDesc
    Dim IGuid As GUID
    
    With tPicConv
       .cbSizeofStruct = Len(tPicConv)
       .picType = vbPicTypeIcon
       .hImage = hIcon
    End With
    
    ' Fill in magic IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    With IGuid
        .data1 = &H7BF80980
        .data2 = &HBF32
        .data3 = &H101A
        .data4(0) = &H8B
        .data4(1) = &HBB
        .data4(2) = &H0
        .data4(3) = &HAA
        .data4(4) = &H0
        .data4(5) = &H30
        .data4(6) = &HC
        .data4(7) = &HAB
    End With
    OleCreatePictureIndirect tPicConv, IGuid, True, oNewPic
    
    Set IconToPicture = oNewPic
    
End Function

'Const EasterEgg As String = "SHOWDEMO"

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
If frmMain.chkOnTop.Checked Then SetTopMostWindow Me.hWnd, True
    'Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.FileDescription
    lblTitle.ToolTipText = App.Title
    lblSite(0).DragIcon = IconToPicture(LoadCursor(0, IDC_HAND))
    lblSite(1).DragIcon = lblSite(0).DragIcon
End Sub

Private Sub lblSite_DragDrop(index As Integer, Source As Control, x As Single, y As Single)
   If Source Is lblSite(index) Then
      lblSite(index).ForeColor = &HFF8080
      Shell "cmd.exe /c START http://" & lblSite(index).Caption, vbHide
   End If
End Sub

Private Sub lblSite_DragOver(index As Integer, Source As Control, x As Single, y As Single, State As Integer)
    If State = vbLeave Then
        With lblSite(index)
            .Drag vbEndDrag
            .Font.Underline = False
        End With
    End If
End Sub

Private Sub lblSite_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
With lblSite(index)
    .Font.Underline = True
    .Drag vbBeginDrag
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not lol Then Exit Sub
frmEE.unl = True
frmEE.Timer1.Enabled = True
lol = False
End Sub

Private Sub picIcon_Click()
'If frmMain.txtURL.Text <> EasterEgg Then Exit Sub
'frmMain.WindowState = vbMinimized
If frmMain.chkOnTop.Checked Then SetTopMostWindow Me.hWnd, False
frmEE.SetFocus
lol = True
End Sub
