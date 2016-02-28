VERSION 5.00
Begin VB.Form frmTesting
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UniBot plugin testing"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtComm 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Optional, for entering a command name."
      Top             =   2420
      Width           =   735
   End
   Begin VB.TextBox txtComm 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   960
      TabIndex        =   3
      ToolTipText     =   "Use | character for splitting parameters (if you're entering more than one)."
      Top             =   2420
      Width           =   2775
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "Execute"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   3615
   End
   Begin VB.CommandButton cmdSettings 
      Caption         =   "Settings"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.ListBox lstLog 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label lblStatus 
      Caption         =   "Idle..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Status"
      Top             =   3120
      Width           =   3615
   End
End
Attribute VB_Name = "frmTesting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Added --

'!!! CAUTION !!!
'IF YOU WANT TO USE PLUGIN UTILITY LATER FOR RESTORING THEN DO NOT MODIFY OR (RE)MOVE ADDED COMMENTS!
'Also note that all between them will be removed in that process, so do not put there anything from the plugin code!

Private Declare Function Testing_SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long 'Added
Private Const LB_SETHORIZONTALEXTENT = &H194
Dim Settings As Object

'-- Added
DOREPLACEHERE
'Added --

Private Sub lstLog_DblClick()
On Error Resume Next
Clipboard.Clear
Clipboard.SetText lstLog.Text
End Sub

Sub addLog(txt As String)
txt = "[" & Now & "] " & txt
If lstLog.ListCount = 32747 Then lstLog.RemoveItem 0
lstLog.AddItem txt
SetListboxScrollbar
lstLog.ListIndex = lstLog.ListCount - 1
lstLog.Text = vbNullString
End Sub

Private Sub SetListboxScrollbar()
Dim new_len As Long
Static max_len As Long
If lstLog.ListCount > 0 Then

        new_len = 10 + lstLog.Parent.ScaleX( _
            lstLog.Parent.TextWidth(lstLog.List(lstLog.ListCount - 1)), _
            lstLog.Parent.ScaleMode, vbPixels)
        If max_len < new_len Then
        max_len = new_len
E:
        Testing_SendMessage lstLog.hWnd, _
        LB_SETHORIZONTALEXTENT, _
        max_len, 0
        End If
Else
max_len = 0
GoTo E
End If
End Sub

Private Sub cmdSettings_Click()
Settings.Show vbModal
End Sub

Private Sub Form_Load()
Me.Show
IPluginInterface_Startup Me
Dim Inf(2) As String
Set Settings = IPluginInterface_Info(Inf)
If Not Settings Is Nothing Then cmdSettings.Enabled = True
cmdExecute_Click 'or: cmdSettings_Click
End Sub

Private Sub cmdExecute_Click()
If txtComm(1).Text = vbNullString Then Exit Sub
Dim Params() As String: Params() = Split(txtComm(0).Text & "|" & txtComm(1).Text, "|")
addLog "Result: " & IPluginInterface_Execute(Params)
MsgBox "Done!", vbInformation
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Settings = Nothing
Class_Terminate
End Sub

'-- Added