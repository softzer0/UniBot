VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UniBot"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   720
   ClientWidth     =   8295
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   8295
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdManager 
      Caption         =   "&..."
      Enabled         =   0   'False
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      ToolTipText     =   "Index manager"
      Top             =   120
      Width           =   375
   End
   Begin prjUniBot.UniTextBox txtOutput 
      Height          =   3015
      Left            =   4200
      TabIndex        =   73
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
   Begin VB.CommandButton cmdR 
      Caption         =   "&R"
      Height          =   255
      Left            =   8400
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   360
      Width           =   135
   End
   Begin VB.CommandButton cmdI 
      Caption         =   "&I"
      Height          =   255
      Left            =   8400
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   120
      Width           =   135
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear log"
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
      Index           =   0
      Left            =   6185
      TabIndex        =   42
      Top             =   120
      Width           =   1985
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save log"
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
      Index           =   0
      Left            =   4200
      TabIndex        =   41
      Top             =   120
      Width           =   1985
   End
   Begin VB.Frame fraS 
      Caption         =   "Strings (extraction)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   67
      Top             =   3000
      Width           =   3975
      Begin VB.VScrollBar VScroll1 
         Enabled         =   0   'False
         Height          =   805
         Index           =   1
         LargeChange     =   1000
         Left            =   3600
         Max             =   100
         SmallChange     =   370
         TabIndex        =   23
         Top             =   410
         Width           =   255
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "+"
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
         Index           =   1
         Left            =   3600
         TabIndex        =   22
         Top             =   150
         Width           =   255
      End
      Begin VB.PictureBox PicBox12 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Index           =   1
         Left            =   120
         ScaleHeight     =   1005
         ScaleWidth      =   3375
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   240
         Width           =   3375
         Begin VB.PictureBox PicBox1 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1000
            Index           =   1
            Left            =   0
            ScaleHeight     =   1005
            ScaleWidth      =   3375
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   0
            Width           =   3375
            Begin VB.CommandButton cmdOpt 
               Caption         =   "..."
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
               Index           =   2
               Left            =   3000
               TabIndex        =   21
               Top             =   720
               Width           =   375
            End
            Begin VB.CommandButton cmdOpt 
               Caption         =   "..."
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
               Index           =   1
               Left            =   3000
               TabIndex        =   18
               Top             =   360
               Width           =   375
            End
            Begin VB.CommandButton cmdOpt 
               Caption         =   "..."
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
               Index           =   0
               Left            =   3000
               TabIndex        =   15
               ToolTipText     =   "String options"
               Top             =   0
               Width           =   375
            End
            Begin VB.TextBox txtExp 
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
               Index           =   2
               Left            =   960
               TabIndex        =   20
               ToolTipText     =   $"frmMain.frx":038A
               Top             =   720
               Width           =   1935
            End
            Begin VB.TextBox txtString 
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
               Index           =   2
               Left            =   0
               TabIndex        =   19
               Top             =   720
               Width           =   855
            End
            Begin VB.TextBox txtExp 
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
               TabIndex        =   17
               ToolTipText     =   $"frmMain.frx":0464
               Top             =   360
               Width           =   1935
            End
            Begin VB.TextBox txtString 
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
               Left            =   0
               TabIndex        =   16
               Top             =   360
               Width           =   855
            End
            Begin VB.TextBox txtExp 
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
               Left            =   960
               TabIndex        =   14
               ToolTipText     =   "Data or/and regular expression(s)"
               Top             =   0
               Width           =   1935
            End
            Begin VB.TextBox txtString 
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
               Left            =   0
               TabIndex        =   13
               ToolTipText     =   "String name"
               Top             =   0
               Width           =   855
            End
            Begin VB.Label lbl2 
               Caption         =   "="
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
               Index           =   2
               Left            =   840
               TabIndex        =   70
               Top             =   720
               Width           =   135
            End
            Begin VB.Label lbl2 
               Caption         =   "="
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
               Index           =   1
               Left            =   840
               TabIndex        =   69
               Top             =   360
               Width           =   135
            End
            Begin VB.Label lbl2 
               Caption         =   "="
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
               Index           =   0
               Left            =   840
               TabIndex        =   68
               Top             =   0
               Width           =   135
            End
         End
      End
   End
   Begin VB.Frame fraR 
      Caption         =   "Request"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   60
      Top             =   480
      Width           =   3975
      Begin VB.TextBox txtURL 
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
         Left            =   600
         TabIndex        =   3
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox txtData 
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
         Left            =   600
         TabIndex        =   4
         ToolTipText     =   "For multipart/form-data: [parameter:value;...]"
         Top             =   600
         Width           =   3255
      End
      Begin VB.Frame fraH 
         Caption         =   "Additional headers"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   61
         Top             =   960
         Width           =   3735
         Begin VB.VScrollBar VScroll1 
            Enabled         =   0   'False
            Height          =   805
            Index           =   0
            LargeChange     =   1000
            Left            =   3360
            Max             =   100
            SmallChange     =   380
            TabIndex        =   12
            Top             =   410
            Width           =   255
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "+"
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
            Index           =   0
            Left            =   3360
            TabIndex        =   11
            ToolTipText     =   "Add one more"
            Top             =   150
            Width           =   255
         End
         Begin VB.PictureBox PicBox12 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1030
            Index           =   0
            Left            =   120
            ScaleHeight     =   1035
            ScaleWidth      =   3135
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   240
            Width           =   3135
            Begin VB.PictureBox PicBox1 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1030
               Index           =   0
               Left            =   0
               ScaleHeight     =   1035
               ScaleWidth      =   3135
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   0
               Width           =   3135
               Begin VB.TextBox txtValue 
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
                  Index           =   2
                  Left            =   1560
                  TabIndex        =   10
                  Top             =   720
                  Width           =   1575
               End
               Begin VB.TextBox txtValue 
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
                  Left            =   1560
                  TabIndex        =   8
                  Top             =   360
                  Width           =   1575
               End
               Begin VB.TextBox txtValue 
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
                  Left            =   1560
                  TabIndex        =   6
                  ToolTipText     =   "Header value"
                  Top             =   0
                  Width           =   1575
               End
               Begin VB.ComboBox cmbField 
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   0
                  ItemData        =   "frmMain.frx":0538
                  Left            =   0
                  List            =   "frmMain.frx":0542
                  TabIndex        =   5
                  ToolTipText     =   "Header name"
                  Top             =   0
                  Width           =   1455
               End
               Begin VB.ComboBox cmbField 
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   1
                  ItemData        =   "frmMain.frx":055A
                  Left            =   0
                  List            =   "frmMain.frx":0564
                  TabIndex        =   7
                  Top             =   360
                  Width           =   1455
               End
               Begin VB.ComboBox cmbField 
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Index           =   2
                  ItemData        =   "frmMain.frx":057C
                  Left            =   0
                  List            =   "frmMain.frx":0586
                  TabIndex        =   9
                  Top             =   720
                  Width           =   1455
               End
               Begin VB.Label lbl1 
                  Caption         =   ":"
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
                  Index           =   2
                  Left            =   1440
                  TabIndex        =   66
                  Top             =   720
                  Width           =   135
               End
               Begin VB.Label lbl1 
                  Caption         =   ":"
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
                  Index           =   1
                  Left            =   1440
                  TabIndex        =   65
                  Top             =   360
                  Width           =   135
               End
               Begin VB.Label lbl1 
                  Caption         =   ":"
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
                  Index           =   0
                  Left            =   1440
                  TabIndex        =   64
                  Top             =   0
                  Width           =   135
               End
            End
         End
      End
      Begin VB.Label Label2 
         Caption         =   "URL:"
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
         TabIndex        =   63
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblD 
         Caption         =   "Post:"
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
         TabIndex        =   62
         ToolTipText     =   "Leave blank for GET type of request."
         Top             =   600
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear output"
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
      Index           =   1
      Left            =   6185
      TabIndex        =   44
      Top             =   6610
      Width           =   1985
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save output"
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
      Index           =   1
      Left            =   4200
      TabIndex        =   43
      Top             =   6610
      Width           =   1985
   End
   Begin VB.TextBox txtName 
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
      Left            =   860
      TabIndex        =   1
      ToolTipText     =   "Name"
      Top             =   120
      Width           =   2760
   End
   Begin VB.ComboBox cmbIndex 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmMain.frx":059E
      Left            =   120
      List            =   "frmMain.frx":05A0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Index"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdProxy 
      Caption         =   "&Proxy && thread settings..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   39
      Tag             =   ","
      Top             =   6960
      Width           =   2775
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
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
      Height          =   375
      Left            =   120
      TabIndex        =   38
      Top             =   6960
      Width           =   5175
   End
   Begin VB.Frame fraE 
      Caption         =   "Else"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   55
      Top             =   6240
      Width           =   3975
      Begin VB.CheckBox chkProxy 
         Caption         =   "Change proxy"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   2520
         TabIndex        =   37
         Top             =   0
         Width           =   1320
      End
      Begin VB.TextBox txtWait 
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
         Left            =   600
         TabIndex        =   34
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cmbGoto 
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
         Height          =   315
         Index           =   1
         ItemData        =   "frmMain.frx":05A2
         Left            =   3145
         List            =   "frmMain.frx":05A9
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   240
         Width           =   785
      End
      Begin VB.ComboBox cmbThen 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         ItemData        =   "frmMain.frx":05B3
         Left            =   2350
         List            =   "frmMain.frx":05BD
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   240
         Width           =   785
      End
      Begin VB.Label lblS 
         Caption         =   "second(s), and:"
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
         Left            =   1170
         TabIndex        =   57
         Top             =   240
         Width           =   1150
      End
      Begin VB.Label lblW 
         Caption         =   "Wait:"
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
         TabIndex        =   56
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame fraT 
      Caption         =   "Then"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   52
      Top             =   5520
      Width           =   3975
      Begin VB.TextBox txtWait 
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
         Left            =   600
         TabIndex        =   30
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox chkProxy 
         Caption         =   "Change proxy"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   2520
         TabIndex        =   33
         Top             =   0
         Width           =   1320
      End
      Begin VB.ComboBox cmbGoto 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         ItemData        =   "frmMain.frx":05CE
         Left            =   3145
         List            =   "frmMain.frx":05D5
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   240
         Width           =   785
      End
      Begin VB.ComboBox cmbThen 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         ItemData        =   "frmMain.frx":05DF
         Left            =   2350
         List            =   "frmMain.frx":05E9
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   240
         Width           =   785
      End
      Begin VB.Label Label7 
         Caption         =   "Wait:"
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
         TabIndex        =   54
         ToolTipText     =   "Alternatively, this will trigger when proxy is not set or enabled."
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "second(s), and:"
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
         Left            =   1170
         TabIndex        =   53
         Top             =   240
         Width           =   1150
      End
   End
   Begin VB.Frame fraI 
      Caption         =   "If"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   51
      Top             =   4440
      Width           =   3975
      Begin VB.CommandButton cmdAdd 
         Caption         =   "+"
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
         Index           =   2
         Left            =   3600
         TabIndex        =   28
         Top             =   600
         Width           =   255
      End
      Begin VB.PictureBox PicBox21 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   3750
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   240
         Width           =   3745
         Begin VB.PictureBox PicBox2 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   310
            Left            =   0
            ScaleHeight     =   315
            ScaleWidth      =   4500
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   0
            Width           =   4500
            Begin VB.ComboBox cmbOper 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               ItemData        =   "frmMain.frx":05FA
               Left            =   3770
               List            =   "frmMain.frx":0604
               Style           =   2  'Dropdown List
               TabIndex        =   27
               Top             =   0
               Width           =   735
            End
            Begin VB.TextBox txtB 
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
               Left            =   2195
               TabIndex        =   26
               Top             =   0
               Width           =   1560
            End
            Begin VB.ComboBox cmbSign 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               ItemData        =   "frmMain.frx":0611
               Left            =   1570
               List            =   "frmMain.frx":061E
               Style           =   2  'Dropdown List
               TabIndex        =   25
               Top             =   0
               Width           =   615
            End
            Begin VB.TextBox txtA 
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
               Left            =   0
               TabIndex        =   24
               ToolTipText     =   "%string%, regular expression, 'text' (use + for joining)"
               Top             =   0
               Width           =   1560
            End
         End
      End
      Begin VB.HScrollBar HScroll1 
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         Max             =   100
         TabIndex        =   29
         Top             =   600
         Width           =   3475
      End
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
      Height          =   3180
      Left            =   4200
      TabIndex        =   40
      Top             =   360
      Width           =   3975
   End
   Begin VB.Timer tmrQ 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7800
      Top             =   7320
   End
   Begin VB.Timer tmrU 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1
      Left            =   6840
      Top             =   7320
   End
   Begin VB.Timer tmrI 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1
      Left            =   7080
      Top             =   7320
   End
   Begin VB.Timer tmrW 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1000
      Left            =   7560
      Top             =   7320
   End
   Begin VB.Label lblStatus 
      Caption         =   "Starting..."
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
      Left            =   720
      TabIndex        =   59
      Top             =   7440
      Width           =   7455
   End
   Begin VB.Label lblS1 
      Caption         =   "Status:"
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
      TabIndex        =   58
      Top             =   7440
      Width           =   495
   End
   Begin VB.Menu conf 
      Caption         =   "Config."
      Begin VB.Menu mnuN 
         Caption         =   "New"
         Begin VB.Menu cmdNew 
            Caption         =   "Blank"
            Enabled         =   0   'False
            Shortcut        =   ^N
         End
         Begin VB.Menu cmdWizard 
            Caption         =   "Wizard..."
            Shortcut        =   ^W
         End
      End
      Begin VB.Menu cmdLoad 
         Caption         =   "Load"
         Shortcut        =   ^L
      End
      Begin VB.Menu cmdSaveC 
         Caption         =   "Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu adv 
      Caption         =   "Advanced"
      Begin VB.Menu chkOnTop 
         Caption         =   "Always on top"
         Shortcut        =   {F1}
      End
      Begin VB.Menu cmdMintoTray 
         Caption         =   "Minimize to tray"
         Shortcut        =   {F2}
      End
      Begin VB.Menu hr 
         Caption         =   "-"
      End
      Begin VB.Menu chkNoSave 
         Caption         =   "Don't save settings on exit"
      End
      Begin VB.Menu cmdAutoSave 
         Caption         =   "Auto-save && exec. Batch..."
      End
      Begin VB.Menu cmdShortcut 
         Caption         =   "Create shortcut..."
         Enabled         =   0   'False
      End
      Begin VB.Menu cmdMake 
         Caption         =   "Make EXE bot..."
         Enabled         =   0   'False
         Shortcut        =   {F5}
      End
      Begin VB.Menu cmdTuning 
         Caption         =   "Fine tuning..."
         Shortcut        =   {F4}
      End
      Begin VB.Menu cmdPlugins 
         Caption         =   "Plugins..."
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu cmdAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const INVALID_HANDLE_VALUE      As Long = -1
Private Const FILE_BEGIN                As Long = &H0
Private Const RT_ICON                   As Long = &H3
Private Const RT_GROUP_ICON             As Long = &HE
Private Const RT_RCDATA = 10&
 
Private Type ICONDIRENTRY
  bWidth          As Byte
  bHeight         As Byte
  bColorCount     As Byte
  bReserved       As Byte
  wPlanes         As Integer
  wBitCount       As Integer
  dwBytesInRes    As Long
  dwImageOffset   As Long
End Type
 
Private Type ICONDIR
  idReserved      As Integer
  idType          As Integer
  idCount         As Integer
End Type
 
Private Type GRPICONDIRENTRY
  bWidth          As Byte
  bHeight         As Byte
  bColorCount     As Byte
  bReserved       As Byte
  wPlanes         As Integer
  wBitCount       As Integer
  dwBytesInRes    As Long
  nID             As Integer
End Type
 
Private Type GRPICONDIR
  idReserved      As Integer
  idType          As Integer
  idCount         As Integer
  idEntries()     As GRPICONDIRENTRY
End Type

Private Declare Function ReadFile Lib "kernel32" (ByVal lFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal lFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function BeginUpdateResource Lib "kernel32" Alias "BeginUpdateResourceA" (ByVal pFileName As String, ByVal bDeleteExistingResources As Long) As Long
Private Declare Function UpdateResource Lib "kernel32" Alias "UpdateResourceA" (ByVal lUpdate As Long, ByVal lpType As Long, ByVal lpName As Long, ByVal wLanguage As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function EndUpdateResource Lib "kernel32" Alias "EndUpdateResourceA" (ByVal lUpdate As Long, ByVal fDiscard As Long) As Long

Private Declare Function PathRelativePathTo Lib "shlwapi.dll" Alias "PathRelativePathToA" (ByVal pszPath As String, ByVal pszFrom As String, ByVal dwAttrFrom As Long, ByVal pszTo As String, ByVal dwAttrTo As Long) As Long
Private Const MAX_PATH As Long = 260
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80

Private WithEvents SystemTray As clsInTray
Attribute SystemTray.VB_VarHelpID = -1

Private Declare Function GetSystemMetrics Lib "user32" ( _
      ByVal nIndex As Long _
   ) As Long

Private Const SM_CXICON = 11
Private Const SM_CYICON = 12

Private Const SM_CXSMICON = 49
Private Const SM_CYSMICON = 50
   
Private Const LR_SHARED = &H8000&

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" ( _
      ByVal hWnd As Long, ByVal wMsg As Long, _
      ByVal wParam As Long, ByVal lParam As Long _
   ) As Long
   
Private Const WM_SETICON = &H80

Private Const ICON_SMALL = 0
Private Const ICON_BIG = 1

Private Const GW_OWNER = 4

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

Private Const CP_UTF16_BE   As Long = 1201       ' UTF16 - big endian.
Private Const CP_UTF32_LE   As Long = 12000      ' UTF32 - little endian.
Private Const CP_UTF32_BE   As Long = 12001      ' UTF32 - big endian.

Const DEFUSERAGENT As String = "Mozilla/5.0 (Windows NT 5.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2490.86 Safari/537.36"

Private Const ProxyRegex = "\b(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\b:\d{2,5}"
Private Const Comms As String = ",rg,rpl,num,dech,dec,enc,u,l,b64,md5,b64d," 'important!
Private WithEvents rh As cAsyncRequests
Attribute rh.VB_VarHelpID = -1
Dim strCmd As String, strPath(1) As String, lngProxyPos() As Long, arrProxy() As String, strNum() As String, lngProxy As Long, bolAb As Boolean, strLastPath1 As String, bolEx As Boolean, bolHid As Boolean, bolSkipLE As VbTriState, bolSilent2 As VbTriState
Dim bytIC As Byte, bolL As Boolean, bolAl As Boolean, bytSh() As Byte, strC As String, bytActive As Byte, intSubT As Integer, intLTmr(1) As Integer, intTmrCount As Integer, bytOrigin As Byte, bolDup As Boolean, strPlg As String, datCompl As Date, strCurrO As String ', strLog As String
Dim bytI As Byte, strName() As String, strURLData() As String, strHeaders() As String, strStrings() As String, strIf() As String, bolProxy() As Boolean, strWait() As String, intGoto() As Integer
Dim colSrc As Collection, colStr As Collection, colPubStr As Collection, colMax As Collection, colInput As Collection, colMaxR As Collection
Public bolDebug As Boolean, strLastPath As String, bolChk As Boolean, strPl As String, lngIF As Long, bytLimit As Byte, bytPlgUse As Byte, bolUnl As Boolean, bolMin As VbTriState

Private Function Subtract(a As String, b As String) As String
    Dim an, bn, rn As Boolean
    Dim ai, bi, barrow As Integer
    an = (Left(a, 1) = "-")
    bn = (Left(b, 1) = "-")
    If an Then a = Mid(a, 2)
    If bn Then b = Mid(b, 2)
    If an And bn Then
        rn = True
    ElseIf bn Then
        Subtract = a
        StrAdd Subtract, b
        Exit Function
    ElseIf an Then
        Subtract = a
        StrAdd Subtract, b
        Subtract = "-" & Subtract
        Exit Function
    Else
        rn = False
    End If
    barrow = Compare(a, b)
    If barrow = 0 Then
        Subtract = "0"
        Exit Function
    ElseIf barrow < 0 Then
        Subtract = a
        a = b
        b = Subtract
        rn = Not rn
    End If
    ai = Len(a)
    bi = Len(b)
    barrow = 0
    Subtract = ""
    Do While ai > 0 And bi > 0
        barrow = barrow + CInt(Mid(a, ai, 1)) - CInt(Mid(b, bi, 1))
        Subtract = CStr(RealMod(barrow, 10)) + Subtract
        barrow = RealDiv(barrow, 10)
        ai = ai - 1
        bi = bi - 1
    Loop
    Do While ai > 0
        barrow = barrow + CInt(Mid(a, ai, 1))
        Subtract = CStr(RealMod(barrow, 10)) + Subtract
        barrow = RealDiv(barrow, 10)
        ai = ai - 1
    Loop
    Do While Len(Subtract) > 1 And Left(Subtract, 1) = "0"
        Subtract = Mid(Subtract, 2)
    Loop
    If Subtract <> "0" And rn Then
        Subtract = "-" + Subtract
    End If
End Function

Private Function RealMod(a As Integer, b As Integer) As Integer
    If a Mod b = 0 Then
        RealMod = 0
    ElseIf a < 0 Then
        RealMod = b + a Mod b
    Else
        RealMod = a Mod b
    End If
End Function

Private Function RealDiv(a As Integer, b As Integer) As Integer
    If a Mod b = 0 Then
        RealDiv = a \ b
    ElseIf a < 0 Then
        RealDiv = a \ b - 1
    Else
        RealDiv = a \ b
    End If
End Function

Private Sub StrAdd(sA As String, sB As String)
    Dim i As Integer
    Dim iLenA As Integer, iLenB As Integer
    Dim iTemp As Integer, iCarry As Integer
    Dim sOut As String
        
    'Calculate once and store, more efficient
    iLenA = Len(sA)
    iLenB = Len(sB)
    
    'Pad out the shorter string with leading 0s
    If (iLenA > iLenB) Then
        sB = String$(iLenA - iLenB, "0") & sB
    ElseIf (iLenB > iLenA) Then
        sA = String$(iLenB - iLenA, "0") & sA
        iLenA = iLenB
    End If
    
    'Allocate string space for the result
    sOut = Space$(iLenA)
    
    For i = iLenA To 1 Step -1
        'Add them together. 48 is the ASCII code for 0, so we subtract
        '96 which is 2*48 to basically convert the returned Asc values
        'into the actual numbers they represent
        iTemp = Asc(Mid$(sA, i, 1)) + Asc(Mid$(sB, i, 1)) + iCarry - 96
        
        'We use Mod 10 so that it only returns the last digit if the number is greater than 9
        Mid$(sOut, i, 1) = iTemp Mod 10
        
        'We use integer divide so that it always rounds down, so 9 \ 10 = 0 instead of 1
        iCarry = iTemp \ 10
    Next i
    
    If iCarry Then sOut = "1" & sOut

    'Return the result
    sA = sOut
End Sub

Private Function Compare(a As String, b As String) As Integer
    Dim an, bn, rn As Boolean
    Dim i, av, bv As Integer
    an = (Left(a, 1) = "-")
    bn = (Left(b, 1) = "-")
    If an Then a = Mid(a, 2)
    If bn Then b = Mid(b, 2)
    If an And bn Then
        rn = True
    ElseIf bn Then
        Compare = 1
        Exit Function
    ElseIf an Then
        Compare = -1
        Exit Function
    Else
        rn = False
    End If
    Do While Len(a) > 1 And Left(a, 1) = "0"
        a = Mid(a, 2)
    Loop
    Do While Len(b) > 1 And Left(b, 1) = "0"
        b = Mid(b, 2)
    Loop
    If Len(a) < Len(b) Then
        Compare = -1
    ElseIf Len(a) > Len(b) Then
        Compare = 1
    Else
        Compare = 0
        For i = 1 To Len(a)
            av = CInt(Mid(a, i, 1))
            bv = CInt(Mid(b, i, 1))
            If av < bv Then
                Compare = -1
                Exit For
            ElseIf av > bv Then
                Compare = 1
                Exit For
            End If
        Next i
    End If
    If rn Then
        Compare = -Compare
    End If
End Function

Function get_relative_path_to(ByVal child_path As String, Optional folder As Boolean, Optional parent_path As String) As String

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
        SendMessage lstLog.hWnd, _
        LB_SETHORIZONTALEXTENT, _
        max_len, 0
        End If
Else
max_len = 0
GoTo E
End If
End Sub

' Purpose:  Take a string whose bytes are in the byte array <the_abytCPString>, with code page <the_nCodePage>, convert to a VB string.
Function FromCPString(ByRef the_abytCPString() As Byte, ByVal the_nCodePage As Long) As String

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

Private Sub SetIcon( _
      ByVal hWnd As Long, _
      ByVal sIconResName As String, _
      Optional ByVal bSetAsAppIcon As Boolean = True _
   )
Dim lhWndTop As Long
Dim lhWnd As Long
Dim cx As Long
Dim cy As Long
Dim hIconLarge As Long
Dim hIconSmall As Long
      
   If (bSetAsAppIcon) Then
      ' Find VB's hidden parent window:
      lhWnd = hWnd
      lhWndTop = lhWnd
      Do While Not (lhWnd = 0)
         lhWnd = GetWindowLong(lhWnd, GW_OWNER)
         If Not (lhWnd = 0) Then
            lhWndTop = lhWnd
         End If
      Loop
   End If
   
   cx = GetSystemMetrics(SM_CXICON)
   cy = GetSystemMetrics(SM_CYICON)
   hIconLarge = LoadImageAsString( _
         App.hInstance, sIconResName, _
         IMAGE_ICON, _
         cx, cy, _
         LR_SHARED)
   If (bSetAsAppIcon) Then
      SendMessageLong lhWndTop, WM_SETICON, ICON_BIG, hIconLarge
   End If
   SendMessageLong hWnd, WM_SETICON, ICON_BIG, hIconLarge
   
   cx = GetSystemMetrics(SM_CXSMICON)
   cy = GetSystemMetrics(SM_CYSMICON)
   hIconSmall = LoadImageAsString( _
         App.hInstance, sIconResName, _
         IMAGE_ICON, _
         cx, cy, _
         LR_SHARED)
   If (bSetAsAppIcon) Then
      SendMessageLong lhWndTop, WM_SETICON, ICON_SMALL, hIconSmall
   End If
   SendMessageLong hWnd, WM_SETICON, ICON_SMALL, hIconSmall
   
End Sub

Private Sub chkNoSave_Click()
chkNoSave.Checked = Not chkNoSave.Checked
End Sub

Private Sub cmdAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub cmdAutoSave_Click()
cmdAutoSave.Caption = Replace(Replace(Replace(Replace(cmdAutoSave.Caption, " (L)", vbNullString), " (O)", vbNullString), " (C)", vbNullString), " (A)", vbNullString)
ShowD
If strPath(0) = vbNullString Then If MsgBox("Continue?", vbQuestion + vbYesNo) = vbNo Then GoTo E
ShowD 1
strCmd = Trim$(InputBox("Enter Batch commands:" & vbLf & "Tips: && = New line; %TFilename%, %TLocation% = File name and location of this program." & vbLf & "Note: Leave blank to disable this option.", "Set Batch commands to execute when process is done", strCmd))
E: PopulA
End Sub

Private Sub ShowD(Optional bytT As Byte)
Dim strPF(1) As String
If strPath(bytT) <> vbNullString Then
strPF(0) = Left$(strPath(bytT), InStrRev(strPath(bytT), "\"))
strPF(1) = Mid$(strPath(bytT), Len(strPF(0)) + 1)
Else: strPF(1) = "{NOW}" 'If bytT = 0 Then strPF(1) = "log" Else: strPF(1) = "results"
End If
Dim strT(1) As String
strT(0) = "Text file (*.txt)|*.txt"
If bytT = 0 Then
strT(0) = "Log file (*.log)|*.log|" & strT(0)
strT(1) = "log (cancel to disable)"
Else
strT(0) = strT(0) & "|HTML file (*.html)|*.html"
strT(1) = "output"
End If
strPath(bytT) = CommDlg(True, "Select where to save " & strT(1), strT(0), strPF(0), Left$(strPF(1), InStr(strPF(1) & ".", ".") - 1))
End Sub

Private Sub PopulA()
If strPath(0) = vbNullString Or strPath(1) = vbNullString Or strCmd = vbNullString Then
If strPath(0) <> vbNullString Then If InStr(cmdAutoSave.Caption, "(L)") = 0 Then cmdAutoSave.Caption = cmdAutoSave.Caption & " (L)"
If strPath(1) <> vbNullString Then If InStr(cmdAutoSave.Caption, "(O)") = 0 Then cmdAutoSave.Caption = cmdAutoSave.Caption & " (O)"
If strCmd <> vbNullString Then If InStr(cmdAutoSave.Caption, "(C)") = 0 Then cmdAutoSave.Caption = cmdAutoSave.Caption & " (C)"
Else: cmdAutoSave.Caption = cmdAutoSave.Caption & " (A)"
End If
cmdAutoSave.Checked = Not Right$(cmdAutoSave.Caption, 1) = "."
End Sub

Sub cmdI_Click()
If Not cmbIndex.Enabled And Me.Enabled Then txtName.SetFocus: Exit Sub
Dim intT As Integer
If Me.Enabled And bytI = cmbIndex.ListCount - 2 Then intT = MsgBox("Do you want to duplicate index?", vbQuestion + vbYesNoCancel)
Select Case True: Case intT = vbCancel, bytI = cmbIndex.ListCount - 1 And Not Filled(cmbIndex.ListCount - 1): cmbIndex.SetFocus: Exit Sub: End Select
If frmManager.bolL Then If frmManager.Tag = "-" Then intT = vbNo: frmManager.Tag = vbNullString
If intT = vbNo Or bytI < cmbIndex.ListCount - 2 Then
AddI
If bolL Then
If frmManager.bolL Then bolL = False: frmManager.Tag = " "
Exit Sub
End If
lblStatus.Caption = "Shifting indexes from " & bytI + 1 & "..."
lblStatus.Refresh
'disF
Screen.MousePointer = 11
Dim i As Byte, j As Byte
i = cmbIndex.ListCount - 1
j = i
Do
j = j - 1
ChngI i, j, True
i = i - 1
Loop Until i = bytI
cmbIndex_Click
cmbIndex.Enabled = False
cmdStart.Enabled = False
EnbC True
EnbIn
Else
E:
If bytI = cmbIndex.ListCount - 1 Then
If Not bolDup Then bolL = False
AddI
If bolL And frmManager.bolL Then bolL = False: frmManager.Tag = " "
bolDup = False
ElseIf Not bolDup Then
If bytI = cmbIndex.ListCount - 2 Then
bytI = bytI + 1
bolDup = True
cmbIndex.ListIndex = bytI
If InStr(cmdOpt(0).Tag, "-" & bytI - 1 & "-") > 0 Then cmdOpt(0).Tag = cmdOpt(0).Tag & "-" & bytI & "-" & Split(Split(cmdOpt(0).Tag, "-" & bytI - 1 & "-")(1), vbLf)(0) & vbLf
GoTo E
End If
Else: bolDup = False
End If
If Me.Enabled Then cmbIndex.SetFocus
End If
End Sub

Private Sub cmdLoad_Click()
Dim strFile As String: strFile = CommDlg(, "Select configuration to load", "Configuration file (*.ini)|*.ini", , "config")
If strFile = vbNullString Then Exit Sub
If Left$(Me.Caption, 1) = "*" Or Not ChkExists Then If Filled(0) Then If MsgBox("Save current configuration?", vbYesNo + vbExclamation) = vbYes Then cmdSaveC_Click
LoadConfig strFile
End Sub

Private Function ChkExists() As Boolean
If CurDir$ <> strInitD Then
strLastPath1 = CurDir$
SetCurrentDirectoryA strInitD
End If
ChkExists = Dir$(cmdSaveC.Tag) <> vbNullString
If strLastPath1 <> vbNullString Then
SetCurrentDirectoryA strLastPath1
strLastPath1 = vbNullString
End If
End Function

Private Function LoadFile2(strPath As String) As Byte()
On Error GoTo E
Open strPath For Binary Access Read As #1
ReDim LoadFile2(LOF(1) - 1)
Get #1, , LoadFile2
Close #1
E:
End Function

Private Sub cmdMake_Click()
If ValidAll Then Exit Sub
Dim i As Byte, a As Byte
For i = 0 To UBound(bolProxy, 2)
For a = 0 To 1
If bolProxy(a, i) And strWait(a, i) = vbNullString Then If MsgBox("You are using option for changing proxy in one or more indexes, and that is not supported in EXE version." & vbNewLine & "Continue anyway?", vbExclamation + vbYesNo) = vbNo Then Exit Sub Else: GoTo E
Next
Next
E: frmEXE.Show vbModal
End Sub

Private Sub cmdManager_Click()
Dim i As Byte
For i = 0 To cmbIndex.ListCount - 1
frmManager.lstI.AddItem i + 1 & vbTab & strName(i)
Next
frmManager.lstI.ListIndex = cmbIndex.ListIndex
SetListboxScrollbar1 frmManager.lstI
frmManager.Show vbModal
End Sub

Private Sub cmdTuning_Click()
frmT.Show vbModal
End Sub

Private Sub LoadConfig(strFile As String)
On Error Resume Next
If lblStatus.Caption <> "Starting..." Then RemC True Else: DimP True
'disF
lblStatus.Caption = "Loading configuration..."
lblStatus.Refresh
Screen.MousePointer = 11
Open strFile For Input Access Read As #1
Dim strL(1) As String, strI As String, i As Byte, j As Integer, a As Integer, intC(1) As Integer, strT As String, strT1 As String, bytT As Byte, s() As String
strI = "|"
While Not EOF(1)
Line Input #1, strL(0)
strL(0) = Trim$(Replace(strL(0), vbCr, vbNullString))
If strL(0) <> vbNullString Then
inp:
If InStr(";#[", Left$(strL(0), 1)) = 0 Then
If InStr(strL(0), "=") > 0 Then
strL(1) = Left$(strL(0), InStr(strL(0), "=") - 1)
If InStr(strI, "|" & strL(1) & "|") = 0 Then
strL(0) = Mid$(Left$(strL(0), Len(strL(0))), Len(strL(1)) + 2)
Select Case strL(1)
Case "url"
If InStr(strI, "|strings|") > 0 Then
If InStr(cmdOpt(0).Tag, "-" & i & "-") > 0 Then
Ad:
strI = "|"
AdN i
End If
ElseIf InStr(strI, "|if|") > 0 Then GoTo Ad
End If
strURLData(i) = strL(0)
If Not bolChk Then
If Not ChkURL(strURLData(i), bolChk) Then
bolChk = Not bolChk
strI = strI & "url|"
End If
ElseIf Not ChkURL(strURLData(i)) Then strI = strI & "url|"
End If
Case "post"
If strL(0) <> vbNullString Then
If Not bolChk Then If Left$(strL(0), 1) = "[" And Right$(strL(0), 1) = "]" And InStr(strL(0), ":") > 0 Then bolChk = Not ChkDat(strL(0)) Else: bolChk = Not ChkStr(strL(0), 1)
strURLData(i) = strURLData(i) & vbLf & strL(0)
strI = strI & "post|"
End If
Case "if"
If InStr(strI, "|strings|") > 0 Then If InStr(cmdOpt(0).Tag, "-" & i & "-") > 0 Then strI = "|": AdN i
If Right$(strL(0), 1) <> """" Then strL(0) = strL(0) & """"
intC(0) = 2
intC(1) = FindC(strL(0), intC(0))
j = 0
strT = vbNullString
Do While Len(strL(0)) > intC(1)
If j > 0 Then
strT = Mid$(strL(0), intC(0), 1)
If Not IsNumeric(strT) Then strT = 0 Else: If strT < 0 Then strT = 0 Else: If strT > 1 Then strT = 1
strT = vbLf & strT
intC(0) = intC(0) + 3
intC(1) = FindC(strL(0), intC(0))
If intC(1) = 0 Then Exit Do
If j > (bytLimit + 1) \ 2 Then AddL True
End If
strIf(i, j) = Trim$(Replace(Mid$(strL(0), intC(0), intC(1) - intC(0)), strC, """"))
If Not bolChk Then bolChk = Not ChkStr(strIf(i, j), 1)
intC(0) = intC(1) + 2
strT1 = Mid$(strL(0), intC(0), 1)
If Not IsNumeric(strT1) Then strT1 = 0 Else: If strT1 < 0 Or strT1 > 5 Then strT1 = 0
strIf(i, j) = strIf(i, j) & vbLf & strT1
intC(0) = intC(0) + 3
intC(1) = FindC(strL(0), intC(0))
If intC(1) = 0 Then strIf(i, j) = vbNullString: Exit Do
strT1 = Trim$(Replace(Mid$(strL(0), intC(0), intC(1) - intC(0)), strC, """"))
If Not bolChk Then bolChk = Not ChkStr(strT1, 1)
If strIf(i, j) = vbNullString And strT1 = vbNullString Then
strIf(i, j) = vbNullString
bytSh(i) = j + 1
If j > 0 Or strT > 0 Then GoTo Ni
Else: strIf(i, j) = strIf(i, j) & vbLf & strT1 & strT
End If
j = j + 1
Ni:
intC(0) = intC(1) + 2
Loop
If j = bytSh(i) Then bytSh(i) = 0
strI = strI & "if|"
Case "strings"
If Right$(strL(0), 1) <> """" Then strL(0) = strL(0) & """"
intC(0) = 2
intC(1) = FindC(strL(0), intC(0))
j = 0
Dim bolS As Boolean
Do While Len(strL(0)) > intC(1)
If Mid$(strL(0), intC(0) - 1, 1) <> """" Then
intC(1) = intC(0) + 7
If Mid$(strL(0), intC(1), 1) <> """" Then Exit Do
strT = vbLf & Mid$(strL(0), intC(0) - 1, intC(1) - intC(0))
s() = Split(strT, ",")
If UBound(s) = 3 Then
If s(1) = "1" Or s(3) = "1" Then bolS = True
Else: strT = vbLf
End If
intC(0) = intC(1) + 1
Else: strT = vbLf
End If
intC(1) = FindC(strL(0), intC(0))
If intC(1) = 0 Then Exit Do
strStrings(i, j) = Trim$(Replace(Replace(Replace(Replace(Replace(Mid$(strL(0), intC(0), intC(1) - intC(0)), strC, """"), "%", ""), "{", ""), "}", ""), "'", ""))
If strStrings(i, j) <> vbNullString Then
For a = 0 To j - 1
If Split(strStrings(i, a), vbLf)(0) = strStrings(i, j) Then strStrings(i, j) = vbNullString: GoTo Ns1
Next
If strT <> vbLf Then If Split(strT, ",")(1) = 1 Then CheckPublic strStrings(i, j), True, True
Else
Ns1:
intC(0) = FindC(strL(0), FindC(strL(0), intC(1) + 3)) + 1
GoTo ns
End If
intC(0) = intC(1) + 3
intC(1) = FindC(strL(0), intC(0))
If intC(1) = 0 Then strStrings(i, j) = vbNullString: Exit Do
strT1 = Trim$(Replace(Mid$(strL(0), intC(0), intC(1) - intC(0)), strC, """"))
If Not bolChk Then bolChk = Not ChkStr(strT1, 1)
intC(0) = intC(1) + 1
If strT1 <> vbNullString Then
strStrings(i, j) = strStrings(i, j) & vbLf & strT1 & strT
If bolS Then
StrAR i, , True
bolS = False
End If
j = j + 1
If j > (bytLimit + 1) \ 2 Then AddL True
Else: strStrings(i, j) = vbNullString
End If
ns:
If Mid$(strL(0), intC(0), 1) <> ";" Then Exit Do
intC(0) = intC(0) + 2
Loop
strI = strI & "strings|"
Case "headers"
If Right$(strL(0), 1) <> """" Then strL(0) = strL(0) & """"
intC(0) = 2
intC(1) = FindC(strL(0), intC(0))
j = 0
Do While Len(strL(0)) > intC(1)
intC(1) = FindC(strL(0), intC(0))
strHeaders(i, j) = Trim$(Replace(Mid$(strL(0), intC(0), intC(1) - intC(0)), strC, """"))
If strHeaders(i, j) <> vbNullString Then
For a = 0 To j - 1
If Split(strHeaders(i, a), vbLf)(0) = strHeaders(i, j) Then strHeaders(i, j) = vbNullString: GoTo Nh1
Next
Else
Nh1:
intC(0) = FindC(strL(0), FindC(strL(0), intC(1) + 3)) + 2
GoTo NH
End If
intC(0) = intC(1) + 3
intC(1) = FindC(strL(0), intC(0))
If intC(1) = 0 Then strHeaders(i, j) = vbNullString: Exit Do
strT1 = Trim$(Replace(Mid$(strL(0), intC(0), intC(1) - intC(0)), strC, """"))
If Not bolChk Then bolChk = Not ChkStr(strT1, 1)
intC(0) = intC(1) + 2
If strT1 <> vbNullString Then
If InStr(fraH.Tag, vbLf & " " & strHeaders(i, j) & " " & vbLf) = 0 Then fraH.Tag = vbLf & " " & strHeaders(i, j) & " " & vbLf & strT1 & fraH.Tag Else: fraH.Tag = Replace(fraH.Tag, vbLf & " " & strHeaders(i, j) & " " & vbLf & Split(Split(fraH.Tag, vbLf & " " & strHeaders(i, j) & " " & vbLf)(1), vbLf)(0) & vbLf, vbLf & " " & strHeaders(i, j) & " " & vbLf & strT1 & vbLf)
strHeaders(i, j) = strHeaders(i, j) & vbLf & strT1
j = j + 1
If j > (bytLimit + 1) \ 2 Then AddL True
Else: strHeaders(i, j) = vbNullString
End If
NH:
If Mid$(strL(0), intC(0), 1) <> """" Then Exit Do
intC(0) = intC(0) + 1
Loop
strI = strI & "headers|"
Case "wait"
If Left$(strL(0), 1) <> """" Then
If Split(strL(0), ";")(0) = vbNullString Then
If strL(0) <> vbNullString Then strWait(0, i) = Val(strL(0))
GoTo N
Else: strWait(0, i) = Val(Split(strL(0), ";")(0))
End If
bytT = Len(strWait(0, i)) + 2
If bytT > Len(strL(0)) Then GoTo N
Else
strWait(0, i) = Mid$(strL(0), 2, FindC(strL(0)) - 2)
If Not bolChk Then bolChk = Not ChkStr(strWait(0, i), 1)
bytT = Len(strWait(0, i)) + 4
If bytT > Len(strL(0)) Then GoTo N
End If
If Mid$(strL(0), bytT, 1) = """" Then
strWait(1, i) = Mid$(strL(0), bytT + 1, FindC(strL(0), bytT + 1) - bytT - 1)
If Not bolChk Then bolChk = Not ChkStr(strWait(1, i), 1)
Else: strWait(1, i) = Val(Mid$(strL(0), bytT))
End If
N:
If strWait(0, i) = "0" Then strWait(0, i) = vbNullString
If strWait(1, i) = "0" Then strWait(1, i) = vbNullString
strI = strI & "wait|"
Case "proxy", "goto"
s() = Split(strL(0) & ";", ";")
If IsNumeric(s(0)) Then If strL(1) = "proxy" Then bolProxy(0, i) = CBool(s(0)) Else: intGoto(0, i) = Abs(CInt(s(0)))
If IsNumeric(s(1)) Then If strL(1) = "proxy" Then bolProxy(1, i) = CBool(s(1)) Else: intGoto(1, i) = Abs(CInt(s(1)))
strI = strI & strL(1) & "|"
End Select
Else
Cl:
strI = "|"
AdN i
GoTo inp
End If
End If
ElseIf Left$(strL(0), 1) = "[" Then
'If InStr(strI, "|name|") = 0 Then
strI = "|" '
AdN i '
strL(0) = Mid$(strL(0), 2, Len(strL(0)) - 2)
If IsNumeric(strL(0)) Then If i + 1 = strL(0) Then GoTo Sk
strName(i) = strL(0)
Sk: 'strI = strI & "name|"
'Else: GoTo Cl
'End If
End If
End If
Wend
Close #1
If Filled(cmbIndex.ListCount - 1) Then If cmbIndex.ListCount <= bytLimit Then AddNew
If cmbIndex.ListCount > 1 Then cmdStart.Enabled = True: EnbC
'If Not Filled(cmbIndex.ListCount - 2) Then
'RemI cmbIndex.ListCount - 2
'cmbIndex.RemoveItem cmbIndex.ListCount - 1
'End If
strT = Mid$(strFile, InStrRev(strFile, "\") + 1)
If InStr(strT, ".") > 0 Then strT = Left$(strT, InStrRev(strT, ".") - 1)
Me.Caption = strT & " - UniBot"
If chkOnTop.Checked Then Me.Caption = Me.Caption & " [on top]"
cmdNew.Enabled = True
cmdSaveC.Tag = strFile
cmbIndex_Click
addLog "Configuration loaded (" & get_relative_path_to(strFile) & ")."
Screen.MousePointer = 0
lblStatus.Caption = "Idle..."
'Exit Sub
'E:
'RemC
'Close #1
'MsgBox "Error in loading configuration!", vbCritical
'If bolDebug Then addLog "Failed to load config. (" & get_relative_path_to(strFile) & ")."
'lblStatus.Caption = "Error! Idle..."
End Sub

Private Sub AdN(i As Byte)
If Not Filled(i) Then
'Dim s() As String
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
bolProxy(a, i) = False
strWait(a, i) = vbNullString
intGoto(a, i) = 0
Next
Else
If cmbIndex.ListCount > bytLimit Then AddL True
AddNew
i = i + 1
End If
End Sub

Private Sub AddNew()
cmbIndex.AddItem cmbIndex.ListCount + 1
cmbGoto(0).AddItem cmbIndex.ListCount
cmbGoto(1).AddItem cmbIndex.ListCount
End Sub

Sub BuildEXE(strLoc As String)
'If cmdStart.Caption = "Stop" Then Exit Sub
frmEXE.Visible = False
lblStatus.Caption = "Building EXE..."
lblStatus.Refresh
Screen.MousePointer = 11
Dim bytT() As Byte
bytT = LoadResData(102, "CUSTOM")
On Error GoTo E
Open strLoc For Binary Access Write As #1
Put #1, , bytT
Close #1
Dim Buf() As Byte, hRes As Long
hRes = BeginUpdateResource(strLoc, 0)
If frmEXE.picIcon(0).Tag <> vbNullString Then ChangeIcon hRes
Dim intL As Integer
Buf() = StrConv(SaveC(intL), vbFromUnicode)
If UpdateResource(hRes, RT_RCDATA, 102&, 0, Buf(0), UBound(Buf) + 1) <> 1 Then GoTo E
Buf() = StrConv(BuildS(intL), vbFromUnicode)
Unload frmEXE
If UpdateResource(hRes, RT_RCDATA, 101&, 0, Buf(0), UBound(Buf) + 1) <> 1 Then GoTo E
If strPlg <> vbNullString Then
Dim lngN As Byte, s() As String, i As Byte
lngN = 103
s() = Split(strPlg, vbLf)
If StrComp(s(0), App.path & IIf(Right$(App.path, 1) <> "\", "\", vbNullString) & "dotnetcomregexlib.dll", vbTextCompare) = 0 Then
ReDim Buf(LOF(2) - 1)
Get #2, , Buf
If UpdateResource(hRes, RT_RCDATA, lngN + i, 0, Buf(0), UBound(Buf) + 1) <> 1 Then GoTo E
i = 1
End If
For i = i To UBound(s()) - 1
Buf() = LoadFile2(s(i))
If UpdateResource(hRes, RT_RCDATA, lngN + i, 0, Buf(0), UBound(Buf) + 1) <> 1 Then GoTo E
Next
End If
EndUpdateResource hRes, 0
strPlg = vbNullString
addLog "EXE bot created (" & get_relative_path_to(strLoc) & ")."
lblStatus.Caption = "Idle..."
Screen.MousePointer = 0
MsgBox "EXE bot has been successfully built!", vbInformation
Exit Sub
E:
On Error Resume Next
If hRes <> 0 Then EndUpdateResource hRes, 0
Kill strLoc
If bolDebug Then addLog "Failed to build EXE (" & get_relative_path_to(strLoc) & ").", True
lblStatus.Caption = "Error! Idle..."
Screen.MousePointer = 0
MsgBox "Error in building EXE bot!", vbCritical
End Sub

Private Function ChangeIcon(ByVal lHandle As Long) As Boolean
  Dim lRet                As Long
  Dim i                   As Integer
  Dim tICONDIR            As ICONDIR
  Dim tGRPICONDIR         As GRPICONDIR
  Dim tICONDIRENTRY()     As ICONDIRENTRY
   
  Dim bIconData()         As Byte
  Dim bGroupIconData()    As Byte
  
  'lngIF = CreateFile(strT, GENERIC_READ, 0, ByVal 0&, OPEN_EXISTING, 0, ByVal 0&)
   
  'If lngIF = INVALID_HANDLE_VALUE Then
  '  ChangeIcon = False
  '  CloseHandle (lngIF)
  '  Exit Function
  'End If
   
  Call ReadFile(lngIF, tICONDIR, Len(tICONDIR), lRet, ByVal 0&)
   
  ReDim tICONDIRENTRY(tICONDIR.idCount - 1)
   
  For i = 0 To tICONDIR.idCount - 1
    Call ReadFile(lngIF, tICONDIRENTRY(i), Len(tICONDIRENTRY(i)), lRet, ByVal 0&)
  Next i
   
  ReDim tGRPICONDIR.idEntries(tICONDIR.idCount - 1)
   
  tGRPICONDIR.idReserved = tICONDIR.idReserved
  tGRPICONDIR.idType = tICONDIR.idType
  tGRPICONDIR.idCount = tICONDIR.idCount
   
  For i = 0 To tGRPICONDIR.idCount - 1
    tGRPICONDIR.idEntries(i).bWidth = tICONDIRENTRY(i).bWidth
    tGRPICONDIR.idEntries(i).bHeight = tICONDIRENTRY(i).bHeight
    tGRPICONDIR.idEntries(i).bColorCount = tICONDIRENTRY(i).bColorCount
    tGRPICONDIR.idEntries(i).bReserved = tICONDIRENTRY(i).bReserved
    tGRPICONDIR.idEntries(i).wPlanes = tICONDIRENTRY(i).wPlanes
    tGRPICONDIR.idEntries(i).wBitCount = tICONDIRENTRY(i).wBitCount
    tGRPICONDIR.idEntries(i).dwBytesInRes = tICONDIRENTRY(i).dwBytesInRes
    tGRPICONDIR.idEntries(i).nID = i + 1
  Next i
   
  'lHandle = BeginUpdateResource(strExePath, False)
  For i = 0 To tICONDIR.idCount - 1
    ReDim bIconData(tICONDIRENTRY(i).dwBytesInRes)
    SetFilePointer lngIF, tICONDIRENTRY(i).dwImageOffset, ByVal 0&, FILE_BEGIN
    Call ReadFile(lngIF, bIconData(0), tICONDIRENTRY(i).dwBytesInRes, lRet, ByVal 0&)
    
    UpdateResource lHandle, RT_ICON, 30001&, 0, Null, 0
    UpdateResource lHandle, RT_ICON, 30002&, 0, Null, 0
    UpdateResource lHandle, RT_ICON, 30003&, 0, Null, 0
    UpdateResource lHandle, RT_GROUP_ICON, 1&, 0, Null, 0
   
    If UpdateResource(lHandle, RT_ICON, tGRPICONDIR.idEntries(i).nID, 0, bIconData(0), tICONDIRENTRY(i).dwBytesInRes) = False Then
      ChangeIcon = False
      CloseHandle (lngIF)
      Exit Function
    End If
     
  Next i
 
  ReDim bGroupIconData(6 + 14 * tGRPICONDIR.idCount)
  CopyMemory ByVal VarPtr(bGroupIconData(0)), ByVal VarPtr(tICONDIR), 6
 
  For i = 0 To tGRPICONDIR.idCount - 1
    CopyMemory ByVal VarPtr(bGroupIconData(6 + 14 * i)), ByVal VarPtr(tGRPICONDIR.idEntries(i).bWidth), 14&
  Next
         
  If UpdateResource(lHandle, RT_GROUP_ICON, 1, 0, ByVal VarPtr(bGroupIconData(0)), UBound(bGroupIconData)) = False Then
    ChangeIcon = False
    CloseHandle (lngIF)
    Exit Function
  End If
   
  'If EndUpdateResource(lHandle, False) = False Then
  '  ChangeIcon = False
  '  CloseHandle (lngIF)
  'End If
 
  Call CloseHandle(lngIF)
  ChangeIcon = True
End Function

Public Sub DetectP()
If cmdStart.Caption = "Stop" Then Exit Sub
lblStatus.Caption = "Checking configuration..."
lblStatus.Refresh
Screen.MousePointer = 11
Dim i As Integer
For i = 0 To cmbIndex.ListCount - 1 + CInt(bytLimit <> cmbIndex.ListCount - 1)
ChkInd i
Next
lblStatus.Caption = "Idle..."
Screen.MousePointer = 0
frmEXE.cmdProceed.Enabled = True
frmEXE.cmdChoose.Enabled = True
If strPl = vbLf Then
If Not bolDebug Then frmEXE.frm1.Enabled = False
frmEXE.lstPlugins.Enabled = False
frmEXE.lbl1.Enabled = False
frmEXE.optTemp.Enabled = False
frmEXE.optCurrent.Enabled = False
'Else: frmEXE.lstPlugins.Enabled = False
End If
End Sub

Private Sub PlugCheck(strExp As String, Optional a As Byte, Optional bytY As Byte, Optional i As Byte, Optional bolT As Boolean)
Dim intC(2) As Long, strT As String, intS As Integer, intP As Integer, bytC As Byte, intL(1) As Long, strT1 As String
strExp = strExp & "+"
intS = 1
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
If InStr(intC(0), strExp, "'") = 0 Then Exit Sub
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
If intC(1) < 2 Then Exit Sub
strT = Mid$(strExp, intC(2), intC(0) - intC(2) - 1)
End If
If InStr(strT, ")") > 0 Then
If intP = 0 Then Exit Sub
If UBound(Split(strT, ")")) < bytC Then bytC = bytC - UBound(Split(strT, ")")) Else: bytC = 0
If InStr(intC(2), strExp, ")") < InStr(intC(2), strExp & "'", "'") Then intC(2) = InStr(intC(2), strExp, ")")
strT = "," & Mid$(strExp, intP, InStr(intP, strExp, "(") - intP) & ","
If InStr(Comms, strT) = 0 Then
If InStr(frmPlugins.strC, strT) > 0 Then
If strT <> ",rg1," Then
strT = Split(frmPlugins.strC, strT)(0)
strT = Mid$(strT, InStrRev(vbLf & strT, vbLf, Len(strT) - 1) + 1)
strT1 = Left$(strT, InStr(strT, "/") - 1)
strT = Mid$(strT, InStr(strT, "/") + 1, Len(strT) - InStr(strT, "/") - 1)
strT1 = Replace(Split(Split(frmPlugins.strP, "|" & strT1 & "|")(1), "|")(0), "S", vbNullString)
If InStr(strPl, vbLf & strT & "|" & strT1 & vbLf) = 0 Then strPl = strPl & strT & "|" & strT1 & vbLf
frmEXE.SetP strT, CByte(strT1)
Else
If InStr(strPl, vbLf & "AdvancedRegEx|0" & vbLf) = 0 Then strPl = strPl & "AdvancedRegEx|0" & vbLf
frmEXE.SetP "AdvancedRegEx", 0
End If
ElseIf bolDebug Then
strT = IIf(bolT, "Warning", "Error") & ": Unknown command " & Replace(strT, ",", """") & " on index " & a + 1 & ", "
Select Case bytY
Case 0: addLog strT & "URL.", True
Case 1: addLog strT & "POST data.", True
Case 2: addLog strT & "header " & i + 1 & ".", True
Case 3: addLog strT & IIf(Not bolT, "crucial ", vbNullString) & "string " & i + 1 & ".", True
Case 4: addLog strT & "If " & i + 1 & ".", True
Case 5: addLog strT & "wait " & i + 1 & ".", True
End Select
End If
End If
bytC = 0
strExp = Replace(strExp, Mid$(strExp, intP, intC(2) - intP + 1), "''")
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
If intC(1) < 2 Then Exit Sub
Else
intC(1) = InStr(intC(0), strExp, "+")
intC(1) = intC(1) - 1
End If
intC(0) = intC(1) + 2
N1:
Loop Until intC(0) >= Len(strExp)
End Sub

Private Function FindC(ByVal strS As String, Optional ByVal intC As Integer = 2) As Integer
FindC = InStr(intC, Left$(strS, intC - 1) & Replace(Mid$(strS, intC), strC, "  "), """")
End Function

Private Sub cmdPlugins_Click()
frmPlugins.Show vbModal
End Sub

Private Sub cmdProxy_Click()
frmPT.Show vbModal
End Sub

Sub cmdR_Click()
Select Case True: Case Not Me.Enabled, bytI <> cmbIndex.ListCount - 1, bytLimit = cmbIndex.ListCount - 1 And Filled(cmbIndex.ListCount - 1): RemI True, Not cmbIndex.Enabled And Me.Enabled: Case Else: cmbIndex.SetFocus: End Select
End Sub

Private Sub cmdSaveC_Click()
If ValidAll Then Exit Sub
If cmdSaveC.Tag = vbNullString Then
F: Dim strFile As String: strFile = CommDlg(True, "Select where to save configuration", "Configuration file (*.ini)|*.ini", , "config")
Else: If ChkExists Then If MsgBox("Overwrite existing configuration file?", vbExclamation + vbYesNo) = vbNo Then GoTo F Else: strFile = cmdSaveC.Tag Else: strFile = cmdSaveC.Tag
End If
If strFile = vbNullString Then Exit Sub
lblStatus.Caption = "Saving cofiguration..."
lblStatus.Refresh
Screen.MousePointer = 11
'disF
On Error GoTo E
Open strFile For Output Access Write As #1
Print #1, ";UniBot configuration file"
Print #1, SaveC;
Close #1
'disF
Dim strT As String
strT = Mid$(strFile, InStrRev(strFile, "\") + 1)
If InStr(strT, ".") > 0 Then strT = Left$(strT, InStrRev(strT, ".") - 1)
Me.Caption = strT & " - UniBot"
If chkOnTop.Checked Then Me.Caption = Me.Caption & " [on top]"
cmdSaveC.Tag = strFile
EnbC
addLog "Configuration saved (" & get_relative_path_to(strFile) & ")."
Screen.MousePointer = 0
lblStatus.Caption = "Idle..."
Exit Sub
E:
Close #1
'disF
If bolDebug Then addLog "Failed to save config. (" & get_relative_path_to(strFile) & ").", True
Screen.MousePointer = 0
lblStatus.Caption = "Error! Idle..."
MsgBox "Error in saving configuration!", vbCritical
End Sub

Private Function SaveC(Optional intE As Integer = -1) As String
Dim s() As String, strT As String, bytS As Byte, i As Byte, a As Byte, bolE As Boolean, bolT(5) As Boolean
If intE = 0 Then
bolE = True
i = cmbIndex.ListCount - 1 + CInt(bytLimit <> cmbIndex.ListCount - 1)
Else: i = cmbIndex.ListCount - 1
End If
For i = 0 To i
If Not bolE Then If strName(i) <> vbNullString Then SaveC = SaveC & vbNewLine & "[" & strName(i) & "]" & vbNewLine Else: SaveC = SaveC & vbNewLine & "[" & i + 1 & "]" & vbNewLine
If Split(strURLData(i) & vbLf, vbLf)(0) <> vbNullString Then SaveC = SaveC & "url=" & Replace(strURLData(i), vbLf, vbNewLine & "post=") & vbNewLine
If strIf(i, 0) <> vbNullString Then
If bolE Then
If Split(strURLData(i) & vbLf, vbLf)(0) = vbNullString Then
If i > 0 Then
If Not bolT(2) Then SaveC = SaveC & "[]" & vbNewLine
bolT(0) = False
End If
Else: bolT(0) = True
End If
bolT(1) = True
bolT(2) = False
End If
For a = 0 To UBound(strIf, 2)
If strIf(i, a) <> vbNullString Then
Dim j As Byte
For j = 1 To bytS
strT = strT & "1,"""",0,"""";"
Next
bytS = 0
s() = Split(strIf(i, a), vbLf)
If a > 0 Then strT = strT & s(3) & ",""" Else: strT = "if="""
strT = strT & Replace(s(0), """", strC) & """," & s(1) & ",""" & Replace(s(2), """", strC) & """;"
ElseIf a > bytSh(i) Then Exit For
Else: bytS = bytS + 1
End If
Next
If bolE Then If a > intE Then intE = a
SaveC = SaveC & Left$(strT, Len(strT) - 1) & vbNewLine
ElseIf bolE Then
bolT(0) = Split(strURLData(i) & vbLf, vbLf)(0) <> vbNullString
bolT(1) = False
End If
If bolE Then
bolT(3) = False
If i > 0 Then If bolT(1) Or bolT(0) Or bolT(2) Then bolT(2) = False Else: bolT(2) = True
End If
If strStrings(i, 0) <> vbNullString Then
If bolE Then If bolT(2) Then SaveC = SaveC & "[]" & vbNewLine Else: bolT(2) = True
strT = "strings="
a = 0
Do While strStrings(i, a) <> vbNullString
s() = Split(strStrings(i, a), vbLf)
If s(2) <> vbNullString Then
strT = strT & s(2) & ":"""
If bolE Then bolT(3) = Split(s(2), ",")(1) = "1" Or Split(s(2), ",")(3) = "1"
Else: strT = strT & """"
End If
strT = strT & Replace(s(0), """", strC) & """,""" & Replace(s(1), """", strC) & """;"
a = a + 1
If UBound(strStrings, 2) < a Then Exit Do
Loop
If bolE Then If a > intE Then intE = a
SaveC = SaveC & Left$(strT, Len(strT) - 1) & vbNewLine
End If
If strHeaders(i, 0) <> vbNullString Then
If bolE Then If Not bolT(0) Then GoTo N
strT = "headers="""
a = 0
Do While strHeaders(i, a) <> vbNullString
s() = Split(strHeaders(i, a), vbLf)
strT = strT & Replace(s(0), """", strC) & """,""" & Replace(s(1), """", strC) & """;"""
a = a + 1
If UBound(strHeaders, 2) < a Then Exit Do
Loop
If bolE Then If a > intE Then intE = a
SaveC = SaveC & Left$(strT, Len(strT) - 2) & vbNewLine
End If
If Not bolE Then
If bolProxy(0, i) Or bolProxy(1, i) Then
SaveC = SaveC & "proxy=" & CInt(bolProxy(0, i)) * (-1)
If bolProxy(1, i) Then SaveC = SaveC & ";" & CInt(bolProxy(1, i)) * (-1) & vbNewLine Else: SaveC = SaveC & vbNewLine
End If
End If
N:
If strWait(0, i) <> vbNullString Or strWait(1, i) <> vbNullString Then
If bolE And i > 0 And Not bolT(4) Then If Not bolT(0) And Not bolT(1) And Not bolT(3) Then SaveC = SaveC & "[]" & vbNewLine: bolT(5) = True
If IsNumeric(strWait(0, i)) Then SaveC = SaveC & "wait=" & CLng(strWait(0, i)) Else: If strWait(0, i) <> vbNullString Then SaveC = SaveC & "wait=""" & strWait(0, i) & """" Else: SaveC = SaveC & "wait=0"
If strWait(1, i) <> vbNullString Then
If IsNumeric(strWait(1, i)) Then
If strWait(1, i) > 0 Then SaveC = SaveC & ";" & CLng(strWait(1, i)) & vbNewLine
Else: SaveC = SaveC & ";""" & strWait(1, i) & """" & vbNewLine
End If
Else: SaveC = SaveC & vbNewLine
End If
If bolE Then bolT(4) = True
ElseIf bolE Then bolT(4) = False
End If
If intGoto(0, i) > 0 Or intGoto(1, i) > 0 Then
If bolE And i > 0 And Not bolT(5) Then If Not bolT(0) And Not bolT(1) And Not bolT(3) Then SaveC = SaveC & "[]" & vbNewLine
SaveC = SaveC & "goto=" & intGoto(0, i)
If intGoto(1, i) > 0 Then SaveC = SaveC & ";" & intGoto(1, i) & vbNewLine Else: SaveC = SaveC & vbNewLine
If bolE Then bolT(5) = True
ElseIf bolE Then bolT(5) = False
End If
Next
If bolE Then intE = intE * 2: If intE < i Then intE = i
End Function

Private Function ValidAll() As Boolean
txtURL_Validate False
If Not cmdSaveC.Enabled And Not cmdMake.Enabled Then GoTo E
txtData_Validate False
Dim i As Integer, C As Boolean, bolT As Boolean
Do
bolT = False
For i = 0 To cmbField.count - 1
'If cmbField(i).Text = vbNullString Or txtValue(i).Text = vbNullString Then Exit For
cmbField_Validate i, C
If C Then
If i > 2 Then VScroll1(0).Value = VScroll1(0).SmallChange * (i - 2)
GoTo E
End If
If i > 0 Then If cmbField(i - 1).Text = vbNullString And txtValue(i - 1).Text = vbNullString And cmbField(i).Text <> vbNullString And txtValue(i).Text <> vbNullString Then bolT = True
Next
Loop Until Not bolT
Do
bolT = False
For i = 0 To txtExp.count - 1
'If txtExp(i).Text = vbNullString Or txtString(i).Text = vbNullString Then Exit For
txtString_Validate i, C
If Not cmdSaveC.Enabled And Not cmdMake.Enabled Then GoTo E
If C Then
If i > 2 Then VScroll1(1).Value = VScroll1(1).SmallChange * (i - 2)
GoTo E
End If
If i > 0 Then If txtString(i - 1).Text = vbNullString And txtExp(i - 1).Text = vbNullString And txtString(i).Text <> vbNullString And txtExp(i).Text <> vbNullString Then bolT = True
Next
Loop Until Not bolT
Dim a As Integer
For i = 0 To txtA.count - 1
If txtA(i).Text = vbNullString And txtB(i).Text = vbNullString Then
If i > 0 Then
If cmbOper(i - 1).ListIndex = 0 Then Exit For
Else: Exit For
End If
End If
a = i
txtA_Validate a, C
If Not cmdSaveC.Enabled And Not cmdMake.Enabled Then GoTo E
If C Then GoTo E
Next
For i = 0 To 1
txtWait_Validate i, C
If C Then GoTo E
Next
Exit Function
E: ValidAll = True
End Function

Private Sub cmdShortcut_Click()
If ValidAll Then Exit Sub
If Left$(frmMain.Caption, 1) = "*" Then cmdSaveC_Click
If Left$(frmMain.Caption, 1) = "*" Then Exit Sub
On Error GoTo E
Dim varT As Variant, strT(1) As String
strT(0) = Mid$(cmdSaveC.Tag, InStrRev(cmdSaveC.Tag, "\") + 1)
If InStr(strT(0), ".") > 0 Then strT(0) = Left$(strT(0), InStrRev(strT(0), ".") - 1)
varT = "UniBot - " & strT(0)
Select Case MsgBox("Put shortcut in Startup folder?", vbQuestion + vbYesNoCancel)
Case vbNo
varT = CommDlg(True, "Select where to save shortcut", "Shortcut file (*.lnk)|*.lnk", , CStr(varT))
If varT = vbNullString Then Exit Sub
strT(0) = Mid$(varT, InStrRev(varT, "\") + 1)
varT = Left(varT, Len(varT) - Len(strT(0)))
Case vbYes
strT(0) = varT & ".lnk"
varT = 7&
Case Else: Exit Sub
End Select
With CreateObject("Shell.Application").NameSpace(varT)
strT(1) = .Self.path & IIf(Right$(.Self.path, 1) <> "\", "\", vbNullString)
If varT = 7 Then
If Dir$(strT(1) & strT(0)) <> vbNullString Then
strT(0) = InputBox("Shortcut with same name already exists. You can name it differently, continue or cancel.", "Shortcut creation", Left$(strT(0), Len(strT(0)) - 4)) & ".lnk"
If strT(0) = ".lnk" Then Exit Sub
End If
End If
Open strT(1) & strT(0) For Output Access Write As #1
Close #1
SetAttr strT(1) & strT(0), vbHidden
Dim strAP As String, strRP As String
strAP = strInitD
If strAP <> strT(1) Then
If Left$(strT(1), 1) = Left$(App.path, 1) Then
strRP = get_relative_path_to(App.path, True, strT(1)) & "\"
If Left$(strRP, 2) <> "\\" Then If MsgBox("Use relative path instead?", vbQuestion + vbYesNo) = vbYes Then strAP = "%CD%"
End If
Else: strAP = "%CD%"
End If
With .Items.Item(CStr(strT(0))).GetLink
.WorkingDirectory = strAP
.Description = App.FileDescription
If strAP = "%CD%" Then
.path = "%windir%\system32\cmd.exe"
.Arguments = "/c start """" """ & strRP & App.EXEName & ".exe" & """ "
strAP = strT(1)
.ShowCommand = 7
.SetIconLocation get_relative_path_to(strInitD, True, Environ$("windir") & "\system32") & "\" & App.EXEName & ".exe", 0
Else
.path = App.path & IIf(Right$(App.path, 1) <> "\", "\", vbNullString) & App.EXEName & ".exe"
.ShowCommand = 1
.SetIconLocation .path, 0
End If
.Arguments = .Arguments & "-c """ & get_relative_path_to(cmdSaveC.Tag, , strAP) & """"
If MsgBox("Start minimized to tray?", vbQuestion + vbYesNo) = vbYes Then .Arguments = .Arguments & " -m"
If cmdAutoSave.Checked Then
If strPath(0) <> vbNullString Then .Arguments = .Arguments & " -l """ & get_relative_path_to(strPath(0), , strAP) & """"
If strPath(1) <> vbNullString Then .Arguments = .Arguments & " -o """ & get_relative_path_to(strPath(1), , strAP) & """"
If strCmd <> vbNullString Then .Arguments = .Arguments & " -e """ & Replace(strCmd, """", strC) & """"
End If
If frmPT.bolNoStartP Or frmPT.bytTimeout <> 20 Or frmPT.strProxy <> vbNullString Or frmPT.bytThreads <> 1 Or frmPT.bytSubThr <> 1 Or frmPT.bolNoRetry Or frmPT.bolNoChange Or frmPT.bytDelay <> 1 Or frmPT.bytMaxR > 0 Or frmPT.bytCycles > 0 Or frmT.intAfter > 0 Then
If MsgBox("Include all settings?", vbQuestion + vbYesNo) = vbYes Then
.Arguments = .Arguments & " -v"
If frmPT.bolNoStartP Then .Arguments = .Arguments & " -w"
If frmT.intAfter > 0 Then .Arguments = .Arguments & " -f" & IIf(frmT.bolHours, "h", vbNullString) & " " & frmT.intAfter
If frmPT.bytTimeout <> 20 Then .Arguments = .Arguments & " -t " & frmPT.bytTimeout
If frmPT.strProxy <> vbNullString Then
If InStr(frmPT.strProxy, vbLf) > 0 Then
Dim strL As String: strL = CommDlg(True, "Select where to save proxy list", "Text file (*.txt)|*.txt|Any file|*.*", , "proxies")
If strL <> vbNullString Then
Open strL For Output Access Write As #1
Print #1, frmPT.strProxy;
Close #1
.Arguments = .Arguments & " -p """ & get_relative_path_to(strL, , strAP) & """"
End If
Else: .Arguments = .Arguments & " -p """ & frmPT.strProxy & """"
End If
If frmPT.bolSame Then .Arguments = .Arguments & " -s"
If frmPT.bolSkip Then .Arguments = .Arguments & " -b"
If Not frmPT.bolNoRetry Then
If Not frmPT.bolNoChange Then
If frmPT.bytCycles > 0 Then .Arguments = .Arguments & " -c " & frmPT.bytCycles
Else: .Arguments = .Arguments & " -g"
End If
GoTo N
Else: .Arguments = .Arguments & " -n"
End If
ElseIf frmPT.bolNoRetry Then .Arguments = .Arguments & " -n"
Else
N:
If frmPT.bytDelay <> 1 Then .Arguments = .Arguments & " -a " & frmPT.bytDelay
If frmPT.bytMaxR > 0 Then .Arguments = .Arguments & " -r " & frmPT.bytMaxR
End If
If frmPT.bytThreads <> 1 Then .Arguments = .Arguments & " -h " & frmPT.bytThreads
If frmPT.bytSubThr <> 1 Then .Arguments = .Arguments & " -u " & frmPT.bytSubThr
End If
End If
.save
End With
End With
SetAttr strT(1) & strT(0), vbNormal
addLog "Shortcut created (" & get_relative_path_to(strT(1) & strT(0)) & ")."
Exit Sub
E:
On Error GoTo -1
On Error Resume Next
SetAttr strT(1) & strT(0), vbNormal
Kill strT(1) & strT(0)
MsgBox "Error: " & err.Description, vbCritical
If bolDebug Then addLog "Failed to create shortcut.", True
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
colInput.add Left$(LTrim$(strT(3)), 1) & strT(2) & vbLf & Mid$(frmInput.strInf, 2), strT(0) & i & Mid$(LTrim$(strT(3)), 2)
If Left$(frmInput.strInf, 1) = "0" Then Exit Function
If Left$(strT(3), 1) <> " " Then strT(3) = " " & strT(3)
t = t + 1
If t > bytT Then Exit Function
strT(0) = Split(strI, vbLf)(1)
GoTo R
End Function

Private Sub cmdStart_Click()
If cmdStart.Caption = "&Start" Then
If ValidAll Then Exit Sub
lblStatus.Caption = "Preparing..."
lblStatus.Refresh
Screen.MousePointer = 11
disF
cmdStart.Enabled = False
Me.Caption = Me.Caption & " (working)"
If cmdMintoTray.Checked And App.LogMode > 0 Then SystemTray.Tip = Me.Caption
rh.Timeout = frmPT.bytTimeout * 1000
Dim bytT As Byte, i As Integer
bytIC = cmbIndex.ListCount - 1 + CInt(bytLimit <> cmbIndex.ListCount - 1)
If frmPT.strProxy <> vbNullString Then
If frmPT.bytThreads = 0 Then
Dim s() As String, intP(1) As Integer
s() = Split(Replace(frmPT.strProxy, vbCr, vbNullString), vbLf)
frmPT.strProxy = vbNullString
For i = 0 To UBound(s())
If s(i) <> vbNullString Then
If RegExpr(ProxyRegex, s(i), , 2) = s(i) Then
If InStr(frmPT.strProxy, s(i)) = 0 Then
intP(0) = intP(0) + 1
frmPT.strProxy = frmPT.strProxy & s(i) & vbNewLine
End If
End If
End If
Next
If bolDebug Then addLog "Found " & intP(0) & " valid proxy addresses.", True
If intP(0) = 0 Then GoTo N
frmPT.strProxy = Left$(frmPT.strProxy, Len(frmPT.strProxy) - 2)
For i = 0 To bytIC
If bolProxy(0, i) And strWait(0, i) = vbNullString Or bolProxy(1, i) And strWait(1, i) = vbNullString Then intP(1) = intP(1) + 1
Next
If intP(1) > 0 Then
If intP(0) > intP(1) Then
bytT = ProcessNumber(intP(0) \ (intP(1) + 1))
If bytT + frmPT.bytSubThr > 255 Then bytT = 255 - frmPT.bytSubThr
ElseIf intP(0) < intP(1) Then
Screen.MousePointer = 0
If strCmd <> vbNullString Then GoTo C
If MsgBox("Insufficient amount of proxies! Continue anyway?", vbExclamation + vbYesNo) = vbYes Then
C:
Screen.MousePointer = 11
bytT = 1
Else: GoTo E
End If
Else: bytT = 1
End If
Else
bytT = ProcessNumber(intP(0))
If bytT + frmPT.bytSubThr > 255 Then bytT = 255 - frmPT.bytSubThr
End If
Else: bytT = frmPT.bytThreads
End If
If frmPT.strProxy <> vbNullString Then
If frmPT.bolSame And bytT > 1 Then ReDim lngProxyPos(bytT - 1)
arrProxy() = Split(frmPT.strProxy, vbNewLine)
lngProxy = UBound(arrProxy) + 1
End If
Else
N: If frmPT.bytThreads = 0 Then bytT = 1 Else: bytT = frmPT.bytThreads
End If
bytT = bytT - 1
intSubT = frmPT.bytSubThr
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
ReDim strNum(bytI)
Dim tmr As Object
If strURLData(0) <> vbNullString Then Set tmr = tmrU Else: Set tmr = tmrI: tmrI(0).Tag = vbNullString
bytOrigin = frmT.bytTOrigin1
If frmT.bolNoEach Then bytOrigin = bytOrigin \ (bytT + 1)
If frmT.intAfter > 0 Then datCompl = DateAdd(IIf(frmT.bolHours, "h", "n"), frmT.intAfter, Now)
bolAb = False
bytActive = 0
If CurDir$ <> strInitD Then
strLastPath1 = CurDir$
SetCurrentDirectoryA strInitD
End If
For i = 0 To bytT
strNum(i) = "0"
If i > 0 Then Load tmr(i)
intTmrCount = intTmrCount + 1
tmr(i).Enabled = True
Next
Set tmr = Nothing
cmdStart.Caption = "&Stop"
cmdStart.Enabled = True
addLog "Process started."
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
ReDim lngProxyPos(0)
Erase arrProxy
intLTmr(0) = 0
intLTmr(1) = 0
cmdProxy.Tag = ","
ReDim strNum(0)
If strLastPath1 <> vbNullString Then
SetCurrentDirectoryA strLastPath1
strLastPath1 = vbNullString
End If
Dim strT As String
Select Case R
Case 0: If Not cmdStart.Enabled Then strT = "aborted" Else: strT = "stopped"
Case 1: strT = "finished"
Case 2: strT = "canceled"
Case 3: strT = "automatically aborted"
End Select
addLog "Process " & strT & "."
lblStatus.Caption = "Idle..."
Screen.MousePointer = 0
Me.Caption = Left$(Me.Caption, InStrRev(Me.Caption, " (working)") - 1) & Mid$(Me.Caption, InStrRev(Me.Caption, " (working)") + 10)
If cmdMintoTray.Checked And App.LogMode > 0 Then SystemTray.Tip = Me.Caption
cmdStart.Caption = "&Start"
disF
If Me.Visible Then cmdStart.SetFocus
If R = 2 Then Exit Sub
If strPath(0) <> vbNullString Then cmdSave_Click 0
If strPath(1) <> vbNullString Then cmdSave_Click 1
If strCmd <> vbNullString Then
'If r = 0 Then If MsgBox("Proceed with executing Batch commands?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
lblStatus.Caption = "Executing Batch commands..."
Shell "cmd.exe /c " & Replace(Replace(strCmd, "%TLocation%", App.path & IIf(Right$(App.path, 1) <> "\", "\", vbNullString)), "%TFilename%", App.EXEName & ".exe"), vbMaximizedFocus
addLog "Batch commands executed."
lblStatus.Caption = "Idle..."
Else
If Not Me.Visible Then SystemTray_MouseUp 1
If (R = 1 Or R = 3) And Not bolDebug And bolMin <> vbTrue Then MsgBox "The process has been finished successfully!", vbInformation
End If
If bolMin = vbTrue Then Unload Me
End Sub

Private Sub SubmitReq(a As Byte, j As Byte, strP As String, Optional O As String, Optional i As Integer, Optional strT1 As String, Optional strTrim As String, Optional strTNum As String)
'del
'strStrings(0, 0) = "string" & vbLf & "1" & vbLf & ",1"
'strStrings(1, 0) = "string" & vbLf & "%string%+4" & vbLf & ",1"
'strStrings(1, 1) = "5" & vbLf & "'haha'" & vbLf & ",1"
'Debug.Print ProceedString("%string%", "", 0, 0, 1, "", "")
'Debug.Print ProceedString("%string% %{%string%}% xxx %{ses}% sad", "", 1, 0, 0, " 1", "")
'Debug.Print ProceedString("%string%+%{%string%}%+'xxx'+%{ses}%+'sad'", "", 1, 0, 0, " 1", "")
'End
'del
Dim strU(1) As String, strD(1) As String, varD As Variant, bolD1 As Boolean, bolT As Boolean, strS As String, strH(1) As String, colH As Collection, Key As String
GetSrc strS, CStr(j), O, a
strU(0) = Split(strURLData(a), vbLf)(0)
If bolAb Then Exit Sub
Dim strT2 As String: If Left$(O, 1) = "/" Or InStr(strT1, "-") > 0 Or InStr(strT1, "+") > 0 Then strT2 = strT1
Dim strCurr As String: strCurr = "{T: " & j + 1 & ", S: " & i & ", I: " & a + 1 & ", O:" & strT2 & "} "
Dim intT As Integer
If i = 0 And Left$(O, 1) <> "/" Then intT = 255 Else: intT = intSubT
If intT - bytActive < 1 Then GoTo E1
If strTrim = vbNullString Then
strTrim = TrimO(O & " " & i)
StrAdd strNum(j), "1"
strTNum = strNum(j)
End If
strU(1) = ProceedString(strU(0), strS, a, j, i, O, strT2, strTrim, strTNum, -2)
If strU(1) = vbNullString Then
If i = 0 Then
addLog "{T: " & j + 1 & ", I: " & a + 1 & ", O:" & strT2 & "} Error: Blank URL!"
GoTo E
Else: GoTo E3
End If
ElseIf strU(0) <> strU(1) Then
If ChkURL(strU(1)) Then
addLog strCurr & "Error: Invalid or blank URL!", True
GoTo E
End If
End If
strD(0) = Split(strURLData(a) & vbLf, vbLf)(1)
If strD(0) <> vbNullString Then bolD1 = True
If intT - bytActive > 0 Then
Dim b As Byte
Do
If bolDebug Then If strU(1) <> strU(0) Then addLog strCurr & "URL: " & strU(1), True
b = 0
Set colH = New Collection
Do While strHeaders(a, b) <> vbNullString
strH(0) = Replace(ProceedString(Split(strHeaders(a, b), vbLf)(0), strS, a, j, i, O, strT2, strTrim, strTNum, -4), vbLf, vbNullString)
If StrPtr(strH(0)) = 0 Then GoTo E
strH(1) = Replace(ProceedString(Split(strHeaders(a, b), vbLf)(1), strS, a, j, i, O, strT2, strTrim, strTNum, -5), vbLf, vbNullString)
If StrPtr(strH(1)) = 0 Then GoTo E
If strH(0) <> vbNullString And strH(1) <> vbNullString Then
If intT - bytActive < 1 Then GoTo E1
colH.add strH(0) & vbLf & strH(1)
If bolDebug Then If strH(0) <> Split(strHeaders(a, b), vbLf)(0) Or strH(1) <> Split(strHeaders(a, b), vbLf)(1) Then addLog strCurr & strH(0) & ": " & strH(1), True
ElseIf bolDebug Then addLog strCurr & "Warning: Unexpected blank add. header at " & b + 1 & ".", True
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
addLog "{T: " & j + 1 & ", I: " & a + 1 & ", O:" & strT2 & "} Warning: File can't be opened or doesn't exist: " & strT(0), True
CatBinaryString bytD, vbCrLf & vbCrLf & vbCrLf
End If
Else: CatBinaryString bytD, "; filename=""""" & vbCrLf & "Content-Type: application/octet-stream" & vbCrLf & vbCrLf
End If
Else
strT(1) = vbCrLf & vbCrLf & ProceedString(strT(0), strS, a, j, i, O, strT2, strTrim, strTNum, -3) & vbCrLf
If Not bolB Then strD(1) = strD(1) & strT(1) Else: CatBinaryString bytD, strT(1)
End If
Else
If strT(1) <> vbNullString Then strT(1) = vbCrLf
strT(1) = strBoundary & vbCrLf & "Content-Disposition: form-data; name=""" & ProceedString(strT(0), strS, a, j, i, O, strT2, strTrim, strTNum, -3) & """"
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
If bolDebug Then If i > 0 Then If VarPtr(varD) <> 0 Then addLog strCurr & "Warning: Unexpected end of pipe (POST).", True
GoTo E
End If
If intT - bytActive < 1 Then Exit Do
If bolDebug Then addLog strCurr & "POST: (multipart/form-data)", True
colH.add "Content-Type" & vbLf & "multipart/form-data; boundary=" & Mid$(strBoundary, 3)
ElseIf Left$(strD(0), 1) = "<" And Right$(strD(0), 1) = ">" Then
strD(0) = Mid$(strD(0), 2, Len(strD(0)) - 2)
On Error GoTo N
If Dir$(strD(0), vbHidden) = vbNullString Then
N:
addLog "{T: " & j + 1 & ", I: " & a + 1 & ", O:" & strT2 & "} Error: File can't be opened or doesn't exist: " & strD(0), True
GoTo E2
Else: varD = LoadFile2(strD(0))
End If
On Error GoTo 0
If intT - bytActive < 1 Then GoTo E1
If bolDebug Then addLog strCurr & "PUT: " & strD(0), True
strU(1) = "*" & strU(1)
Else
If strD(0) <> "''" Then varD = ProceedString(strD(0), strS, a, j, i, O, strT2, strTrim, strTNum, -3) Else: varD = vbNullChar
If intT - bytActive < 1 Then Exit Do
If bolDebug Then If strD(0) <> varD Then addLog strCurr & "POST: " & Replace(varD, vbLf, "[nl]"), True
If Left$(strD(0), 1) = "{" And Right$(strD(0), 1) = "}" Then colH.add "Content-Type" & vbLf & "application/json" Else: colH.add "Content-Type" & vbLf & "application/x-www-form-urlencoded"
End If
End If
'Dim strT3 As String: If strT1 = vbNullString Then strT3 = IIf(strT2 <> vbNullString Or i > 0, o, vbNullString)
Key = j & "," & i & "," & a & "," & O & "," & strP & "," & strT1 & "," & strTrim & "," & strTNum
If intT - bytActive < 1 Then Exit Do
On Error GoTo err
rh.AddRequest(Key).SendRequest strU(1), varD, colH, strP
On Error GoTo 0
bytActive = bytActive + 1
If tmrQ.Enabled Then tmrQ.Tag = Replace(tmrQ.Tag, "-" & a & "," & j & "," & O & "," & strTrim & "," & strTNum & "," & strP & "," & i & "," & strT1 & vbLf, vbNullString, , 1)
addLog strCurr & "Request sent."
bytD = ""
Set colH = Nothing
i = i + 1
strCurr = Replace(strCurr, "S: " & i - 1 & ",", "S: " & i & ",")
If i > Val(PrepareCol(colMax, a & "," & j & "," & O)) Then Exit Sub
'Debug.Print i, Val(PrepareCol(colMax, a & "," & j & "," & O)), j, O
If i = 1 Then intT = intSubT
strTrim = Left$(strTrim, InStrRev(strTrim, " ")) & i
StrAdd strNum(j), "1"
strTNum = strNum(j)
If intT - bytActive < 1 Then Exit Do
strU(1) = ProceedString(strU(0), strS, a, j, i, O, strT2, strTrim, strTNum, -2)
If strU(1) = vbNullString Then
E3:
If bolDebug Then If StrPtr(strU(1)) <> 0 Then addLog strCurr & "Warning: Unexpected end of pipe (URL).", True
GoTo E
End If
Loop Until intT - bytActive < 1
End If
E1:
If Not bolAb Then
If InStr(tmrQ.Tag, "-" & a & "," & j & "," & O & "," & strTrim & "," & strTNum & "," & strP & "," & i & "," & strT1 & vbLf) = 0 Then
'If Not strU(1) = vbNullString Then 'And Not bolD1 Or bolD1 And Not strU(1) = vbNullString And Not varD = vbNullString
tmrQ.Tag = tmrQ.Tag & "-" & a & "," & j & "," & O & "," & strTrim & "," & strTNum & "," & strP & "," & i & "," & strT1 & vbLf
If bolDebug Then addLog strCurr & "Added to queue.", True
If Not tmrQ.Enabled Then tmrQ.Enabled = True
'End If
End If
End If
GoTo E2
E: If tmrQ.Enabled Then If InStr(tmrQ.Tag, "-" & a & "," & j & "," & O & "," & strTrim & "," & strTNum & "," & strP & "," & i & "," & strT1 & vbLf) > 0 Then tmrQ.Tag = Replace(tmrQ.Tag, "-" & a & "," & j & "," & O & "," & strTrim & "," & strTNum & "," & strP & "," & i & "," & strT1 & vbLf, vbNullString, , 1): Exit Sub
E2:
PrepTr CStr(j), CStr(i), CStr(a), O, strTNum, True
If bolUnl Then Exit Sub
If Not tmrQ.Enabled Then ChkT Else: ChkSt
Exit Sub
err:
If err.Number <> 457 Then
rh.RemoveRequest Key
addLog strCurr & "Error: " & err.Description & ".", True
GoTo E
Else: GoTo E1
End If
End Sub

Private Sub GetSrc(strS As String, j As String, ByVal O As String, a As Byte)
If bytOrigin > 0 Then
Dim strT(3) As String, intT1 As Integer, i As Integer, intT(1) As Integer, bolT2(3) As Boolean
O = Replace(Replace(O, "|", " "), "/", vbNullString, , 1) & " "
strT(3) = Replace(StrReverse(O), "...", vbNullString, , 1)
intT1 = Len(strT(3)) - 1
bolT2(0) = Left$(O, 1) = "/"
bolT2(1) = Mid$(O, 1 + CInt(bolT2(0)) * -1, 3) = "..."
For i = colSrc.count To 1 Step -1
strT(1) = colSrc.Item(i)
If Left$(strT(1), Len(j) + 1) = j & "," Then
strT(1) = Mid$(strT(1), Len(j) + 2)
strT(1) = Left$(strT(1), InStr(strT(1), "\") - 1)
strT(2) = Replace(Replace(strT(1), "|", " "), "/", vbNullString, , 1)
If InStr(O, strT(2) & " ") = 1 Then
intT(0) = Len(O) - Len(strT(2)) - 1
ElseIf bolT2(1) Xor bolT2(3) Then
If InStr(strT(3), " " & Left$(StrReverse(Replace(strT(2), "...", vbNullString, , 1)), intT1)) = 1 Then intT(0) = intT1 - Len(Replace(strT(2), "...", vbNullString, , 1)) Else: intT(0) = -1
Else: If bolT2(1) Or InStr(strT(2), "...") > 0 Then intT(0) = -Len(strT(1)) Else: intT(0) = -1
End If
If intT(0) <> -1 Then
''C:
If strT(0) <> vbNullString Then
If intT(0) > 0 Then
Select Case True
Case intT(1) < 0, intT(0) < intT(1): GoTo G
End Select
ElseIf intT(0) < 0 Then
If intT(1) < 0 Then If bolT2(0) Xor bolT2(2) Or intT(0) >= intT(1) Then GoTo G
Else: GoTo G
End If
Else
G:
strT(0) = strT(1)
strS = Mid$(colSrc.Item(i), Len(strT(1)) + Len(j) + 3)
intT(1) = InStr(strT(1), ",") + 1
bolT2(2) = Mid$(strT(1), intT(1), 1) = "/"
bolT2(3) = Mid$(strT(1), intT(1) + CInt(bolT2(2)) * -1, 3) = "..."
If intT(0) = 0 Then Exit For Else: intT(1) = intT(0) 'Not frmT.bolColl
End If
''ElseIf Replace(strT(1), "/", vbNullString, , 1) = j & ",..." Then GoTo C
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

Private Sub cmdWizard_Click()
frmWizard.Show vbModal
End Sub

Private Sub rh_ResponseFinished(Req As cAsyncRequest)
If bolAb Then Exit Sub
Dim s() As String, strS As String
strS = Req.http.Status & " " & Req.http.StatusText & vbNewLine & Req.http.GetAllResponseHeaders
On Error Resume Next
strS = strS & FromCPString(Req.http.ResponseBody, CP_UTF8) 'strS = strS & Req.http.ResponseText
'If Err Then strS = strS & FromCPString(Req.http.ResponseBody, CP_UTF8)
On Error GoTo 0
s() = Split(Req.Key, ",")
rh.RemoveRequest Req
bytActive = bytActive - 1
If bolDebug Then
Dim strT As String: If Left$(s(3), 1) = "/" Or InStr(s(5), "-") > 0 Or InStr(s(5), "+") > 0 Then strT = s(5)
addLog "{T: " & s(0) + 1 & ", S: " & s(1) & ", I: " & s(2) + 1 & ", O:" & strT & "} Response status - " & Split(strS, vbNewLine)(0), True
End If
CheckIf s(0), s(1), s(2), s(3), s(4), strS, s(5), s(6), s(7)
End Sub

Private Sub rh_Error(Req As cAsyncRequest, ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
rh.RemoveRequest Req
bytActive = bytActive - 1
If bolAb Then Exit Sub
Dim s() As String: s() = Split(Req.Key, ",")
Dim strT(1) As String, strP As String
If Left$(s(3), 1) = "/" Or InStr(s(5), "-") > 0 Or InStr(s(5), "+") > 0 Then strT(1) = s(5)
strT(0) = "T: " & s(0) + 1 & ", S: " & s(1) & ", I: " & s(2) + 1 & ", O:" & strT(1)
addLog "{" & strT(0) & "} Error: " & ErrorDescription & "."
Dim bytT As Byte
If UBound(lngProxyPos) = 0 Then bytT = 0 Else: bytT = s(0)
If Not frmPT.bolNoRetry And lngProxyPos(bytT) > -1 Then
If frmPT.bytMaxR > 0 Then
Dim strT1 As String, i As Byte
strT1 = PrepareCol(colMaxR, s(0) & "," & s(3)) & ","
i = Val(Split(strT1, ",")(0)) + 1
On Error Resume Next
colMaxR.Remove s(0) & "," & s(3)
On Error GoTo 0
If i > frmPT.bytMaxR Then
N1:
If Not frmPT.bolNoChange Then
If frmPT.bytMaxR = 0 Then GoTo N
i = Val(Split(strT1, ",")(1)) + 1
If frmPT.bytCycles > 0 Then
If i <= frmPT.bytCycles Then
If i < frmPT.bytCycles Then colMaxR.add "0," & i, s(0) & "," & s(3)
N:
If frmPT.strProxy <> vbNullString Then
If frmPT.bolSkip Then cmdProxy.Tag = cmdProxy.Tag & s(4) & ","
If SetProxy(strP, CByte(s(0)), strT(0)) Then
lngProxyPos(bytT) = -1
If frmPT.bolSame Then addLog "{T: " & s(0) & "} Error: Out of working/valid proxies!" Else: addLog "Error: Out of working/valid proxies!"
GoTo E
End If
End If
Else: GoTo E
End If
Else: GoTo N
End If
ElseIf frmPT.bytMaxR > 0 Then GoTo E
End If
Else: colMaxR.add i & "," & Val(Split(strT1, ",")(1)), s(0) & "," & s(3)
End If
Else: GoTo N1
End If
If frmPT.bytDelay > 0 Then
If bolDebug Then addLog "{" & strT(0) & "} Waiting " & frmPT.bytDelay & " second(s) before another retry...", True
If bolAb Or ChkSt Then Exit Sub
SubmitIW frmPT.bytDelay & "," & s(2) & "," & s(0) & "," & strP & "," & s(3) & "," & s(1) & "," & s(5) & "," & s(6) & ",,,," & vbNullChar
Else
If strP = vbNullString Then strP = s(4)
SubmitReq CInt(s(2)), CByte(s(0)), strP, s(3), CInt(s(1)), s(5), s(6), s(7)
End If
Exit Sub
End If
E:
PrepTr s(0), s(1), s(2), s(3), s(7), True
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
If frmT.intAfter = 0 Then Exit Function
If Now < datCompl Then Exit Function
If Not bolT Then Enb 3
End If
ChkSt = True
End Function

Private Function SetProxy(strP As String, bytI As Byte, strT As String) As Boolean
If bolAb Then Exit Function
Dim bytT As Byte, lngP As Long
If UBound(lngProxyPos) = 0 Then bytT = 0 Else: bytT = bytI
If lngProxyPos(bytT) = lngProxy Then
lngP = 0
R:
lngProxyPos(bytT) = 0
Else: lngP = lngProxyPos(bytT)
End If
If bolAb Then Exit Function
If frmPT.bytThreads > 0 Then
R1:
Do
If lngProxyPos(bytT) = lngProxy Then If lngP > 0 Then GoTo R Else: GoTo E
SetP strP, lngProxyPos(bytT)
If lngProxyPos(bytT) = lngP Then
E:
SetProxy = True
Exit Function
End If
Loop Until RegExpr(ProxyRegex, strP, , 2) = strP Or bolAb
If Not bolAb Then If frmPT.bolSkip Then If InStr(cmdProxy.Tag, "," & strP & ",") > 0 Then GoTo R1
Else: SetP strP, lngProxyPos(bytT)
End If
If bolDebug Then addLog "{" & strT & "} Proxy (" & lngProxyPos(bytT) & "): " & strP, True
End Function

Private Sub SetP(strP As String, lngP As Long)
strP = arrProxy(lngP)
lngP = lngP + 1
End Sub

Private Function ProceedString(ByVal strInp As String, strS As String, bytI As Byte, j As Byte, i As Integer, O As String, o2 As String, strTrim As String, strTNum As String, Optional intI As Integer, Optional strI As String, Optional bolA As Boolean, Optional bytA As Byte) As String
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
If intC(1) = 0 Or intC(0) > intC(1) Then Exit Do 'If intC(0) > intC(1) Then GoTo R
strT(0) = Mid$(strInp, intC(0), intC(1) - intC(0))
If bolT1 Then
bolT1 = False
If strT(0) <> vbNullString Then If i = 0 Then strT(2) = ProceedString(strT(0), strS, bytI, j, i, O, o2, strTrim, strTNum, intT) Else: strT(2) = ProceedString(strT(0), strS, bytI, j, i, O, o2, strTrim, strTNum)
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
strT(1) = PrepareCol(colStr, strT(0) & "," & bytI & "," & j & "," & i & "," & strTNum)
If StrPtr(strT(1)) <> 0 Then strT(1) = Mid$(strT(1), InStr(strT(1), "\") + 1)
If StrPtr(strT(1)) = 0 Or Replace(strI, "%", vbNullString, , 1) = strT(0) Then
If Replace(strI, "%", vbNullString, , 1) <> strT(0) Or Not bolA And i = 0 Then
If Left$(strI, 1) <> "%" Then
If bytA > 0 Then a = bytA - 1 Else: a = 0
Do While strStrings(bytI, a) <> vbNullString
If Left$(strStrings(bytI, a), InStr(strStrings(bytI, a), vbLf) - 1) = strT(0) Then
'strT(1) = Split(strStrings(bytI, a), vbLf)(1)
GoTo C1
End If
a = a + 1
If UBound(strStrings, 2) < a Then Exit Do
Loop
End If
Dim strI1(6) As String, strI2(1) As String, intT2 As Integer 'strI1(2), intT2(1), bolT2(3) As Boolean
bolT1 = False
strI1(5) = "0"
strT(1) = vbNullString
If strI2(0) <> vbNullString Then
strI2(0) = vbNullString
strI2(1) = vbNullString
End If
strT2(0) = j & "," & Replace(Replace(O, "|", " "), "/", vbNullString, , 1) & " " 'Replace(Replace(O, "|", " "), "/", vbNullString, , 1) & " " & i
strI1(3) = Replace(StrReverse(strT2(0)), "...", vbNullString, , 1)
intT2 = Len(strI1(3)) - 1
'strI1(3) = StrReverse(Replace(Mid$(strT2(0), InStr(strT2(0), ",") + 1), "...", vbNullString, , 1))
'bolT2(0) = Left$(O, 1) = "/"
'bolT2(1) = Mid$(O, 1 + CInt(bolT2(0)) * -1, 3) = "..."
'R1:
'to be fixed: START
For intT1 = colPubStr.count To 1 Step -1
strI1(0) = colPubStr.Item(intT1)
If Left$(strI1(0), InStr(strI1(0), "%") - 1) = strT(0) Then
strI1(0) = Mid$(strI1(0), Len(strT(0)) + 2)
strI1(0) = Left$(strI1(0), InStr(strI1(0), "{") - 1)
strI1(4) = Subtract(strTNum, Left$(strI1(0), InStr(strI1(0), ",") - 1))
If Left$(strI1(4), 1) <> "-" Then
strI1(6) = Mid$(strI1(0), InStr(strI1(0), ",") + 1)
If strI1(5) <> "0" Then
If Not bolT1 Then If strI1(6) <> bytI Then GoTo N2
If Compare(strI1(5), strI1(4)) <> -1 Then GoTo N2
End If
strI1(1) = Mid$(colPubStr.Item(intT1), Len(strT(0)) + Len(strI1(0)) + 3)
strI1(1) = Left$(strI1(1), InStr(InStr(strI1(1), ",") + 1, strI1(1), "\") - 1)
strI1(2) = Replace(Replace(strI1(1), "|", " "), "/", vbNullString, , 1)
Select Case True
Case InStr(strT2(0), strI1(2) & " ") = 1, InStr(strI1(3), " " & Left$(StrReverse(Replace(strI1(2), "...", vbNullString, , 1)), intT2)) = 1
'intT2(0) = Len(strT2(0)) - Len(strI1(2))
'ElseIf bolT2(1) Xor bolT2(3) Then
'If InStr(" " & strT2(0), " " & Left$(StrReverse(Replace(strI1(2), "...", vbNullString, , 1)), Len(strT2(0)))) = 1 Then intT2(0) = Len(strT2(0)) - Len(Replace(strI1(2), "...", vbNullString, , 1)) Else: intT2(0) = -1
'Else: If bolT2(1) Or InStr(strI1(2), "...") > 0 Then intT2(0) = -Len(strI1(2)) Else: intT2(0) = -1
'End If
'If intT2(0) <> -1 Then
'C2:
'If strT(1) <> vbNullString Then
'If Not bolT1 Then If Replace(strI, "%", vbNullString, , 1) <> strT(0) Or strI1(0) <> bytI Then GoTo N2 'And bytI > strI2(0)
'If intT2(0) > 0 Then
'Select Case True
'Case intT2(1) < 0, intT2(0) < intT2(1): GoTo G
'End Select
'ElseIf intT2(0) < 0 Then
'If intT2(1) < 0 Then If bolT2(0) Xor bolT2(2) Or intT2(0) >= intT2(1) Then GoTo G
'Else: GoTo G
'End If
'Else
'G:
bolT1 = strI1(6) <> bytI
strI2(0) = strI1(6)
strI2(1) = strI1(1)
'bolT2(2) = Mid$(strI2(1), InStr(strI2(1), ",") + 1, 1) = "/"
'bolT2(3) = Mid$(strI2(1), 3 + CInt(bolT2(2)) * -1, 3) = "..."
strT(1) = Mid$(colPubStr.Item(intT1), Len(strT(0)) + Len(strI1(0)) + Len(strI1(1)) + 4)
If strI1(4) = 0 Or Replace(strI, "%", vbNullString, , 1) <> strT(0) Then Exit For Else: strI1(5) = strI1(4) 'Exit For 'Not frmT.bolColl
'End If
''ElseIf Replace(strI1(1), "/", vbNullString, , 1) = j & ",..." Then GoTo C2
End Select
End If
End If
N2:
If bolAb Then Exit Function
Next
bolT1 = False
'to be fixed: END
'If strT(0) = "id" Then
'Open "problem.txt" For Append As #1
'Print #1, "CHOSEN: " & strT(1)
'Print #1, "COUNT: " & colPubStr.count, "ORIGIN: " & O, "CURR. NUMBER: " & strTNum
'For intT1 = 1 To colPubStr.count
'Print #1, colPubStr.Item(intT1)
'Next
'Print #1,
'Close #1
'End If
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
'ElseIf strT2(0) = j & "," & Replace(O, "/", vbNullString, , 1) & " " & i Then
'strT2(1) = strT2(0)
'strT2(0) = j & "," & Replace(Replace(strTrim, "|", " "), "/", vbNullString, , 1)
'If strT2(0) <> strT2(1) Then GoTo R1
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
If strT(1) = vbNullString Or intT1 = 0 Then strT2(0) = ProceedString(strT(1), strS, bytI, j, 0, O, o2, strTrim, strTNum, intT, "%" & strT(0), True)
Else: strT2(0) = ProceedString(strT(1), strS, bytI, j, i, O, o2, strTrim, strTNum, , "%" & strT(0))
End If
Else: strT2(0) = ProceedString(strT(1), strS, bytI, j, i, O, o2, strTrim, strTNum, -1, "%" & strT(0))
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
addLog "{T: " & j + 1 & ", S: " & i & ", I: " & bytI + 1 & ", O:" & o2 & "} Error: Crucial string %" & strT(0) & "% is blank."
bolT1 = True
GoTo Ex
End If
If strT(1) <> strT2(0) Then
If bolDebug Then If intT = -1 Or intT = 0 Then addLog "{T: " & j + 1 & ", S: " & i & ", I: " & bytI + 1 & ", O:" & o2 & "} %" & strT(0) & "%: " & strT2(0), True
strT(1) = strT2(0)
End If
If bolAb Then Exit Function
If strT(1) <> vbNullString And intT1 > 0 Then GoTo Ex
If s(1) = "1" Then
'If s(2) = "1" Then
'If i = 0 Then
'strT2(0) = " " & i
'strT2(1) = "," & strT(1)
'Else
'strT2(0) = vbNullString
'strT2(1) = i & "," & strT(1)
'End If
'Else
'strT2(0) = " " & i
'strT2(1) = "," & strT(1)
'End If
'Dim bolT1 As Boolean
'On Error GoTo N1
'Do
'o1 = Left$(o1, InStrRev(o1, " ") - 1)
'colPubStr.Remove strT(0) & "," & J & "," & o1
'bolT1 = True
'C:
'Loop Until o1 = vbNullString Or bolAb
'N2:
'On Error GoTo -1
'If bytOrigin = 0 Then
'If bytA > 0 Or intI < -5 Then
If s(2) <> "1" Or i > 0 Then strT2(0) = strTrim Else: strT2(0) = O
If strT2(0) <> O Then strT2(1) = "," & strT(1) Else: strT2(1) = i & "," & strT(1)
'Else
'strT2(0) = O
'strT2(1) = "," & strT(1)
'End If
'If StrPtr(PrepareCol(colPubStr, strT(0) & "," & j & "," & O & strT2(0))) <> 0 Then Return
'Debug.Print strT2(0), strT2(1)
On Error Resume Next
colPubStr.Remove strT(0) & "%" & j & "," & strT2(0) & "," & strTNum
On Error GoTo 0
'End If
colPubStr.add strT(0) & "%" & strTNum & "," & bytI & "{" & j & "," & strT2(0) & "\" & intT & "\" & strT2(1), strT(0) & "%" & j & "," & strTNum
End If
colStr.add j & "," & strTNum & "\" & strT(1), strT(0) & "," & j & "," & strTNum
If s(3) = "1" Then
If frmT.intOutMax > 0 Then If (Len(txtOutput.Text) - Len(Replace(txtOutput.Text, vbNewLine, vbNullString))) / 2 = frmT.intOutMax Then txtOutput.Text = Mid$(txtOutput.Text, InStr(txtOutput.Text, vbNewLine) + 2)
txtOutput.Text = txtOutput.Text & Replace(Replace(Replace(Replace(Replace(Replace(frmT.strTemplate0, "{T}", j + 1), "{S}", i), "{O}", o2), "{I}", bytI + 1), "{N}", strT(0)), "{D}", Now) & strT(1) & frmT.strTemplate1
txtOutput.ScrollToBottom
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
colMax.add intI - 1, bytI & "," & j & "," & O
End Function

Private Sub CheckIf(s0 As String, s1 As String, s2 As String, s3 As String, s4 As String, strS As String, strT1 As String, strTrim As String, strTNum As String, Optional bolEn As Boolean)
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
con(0) = ProceedString(s(0), strS, CByte(s2), CByte(s0), CInt(s1), s3, strT2, strTrim, strTNum, -6)
con(1) = ProceedString(s(2), strS, CByte(s2), CByte(s0), CInt(s1), s3, strT2, strTrim, strTNum, -7)
If bolAb Then Exit Sub
If bolDebug Then
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
'If sp(2) = "1" And s1 = "0" Then strT(0) = s3 Else: strT(0) = s3 & " " & s1
If PrepareCol(colStr, s(0) & "," & s0 & "," & strTNum) = vbNullString And PrepareCol(colPubStr, s(0) & "%" & s0 & "," & strTNum) = vbNullString Then
If StrPtr(ProceedString("%" & s(0) & "%", strS, CByte(s2), CByte(s0), CInt(s1), s3, strT2, strTrim, strTNum, , , , i + 1)) = 0 Then GoTo E
'If ExtrStr(s(0), s(1), Replace(s(0), ",", ",,"), strS, i, CByte(s2), CByte(s0), CInt(s1), s3, strT2) Then
'End If
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
PrepTr s0, s1, s2, s3, strTNum
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
If Not bolProxy(ns(0), s2) Then
If strWait(ns(0), s2) <> vbNullString Then
C:
If bolAb Or ChkSt Then Exit Sub
strT(0) = ProceedString(strWait(ns(0), s2), strS, CByte(s2), CByte(s0), CInt(s1), s3, strT2, strTrim, strTNum, -8)
If IsNumeric(strT(0)) Then
If bolDebug Then
If ns(0) = 0 Then strT(1) = "Then" Else: strT(1) = "Else"
If strT(0) <> strWait(ns(0), s2) Then addLog "{T: " & s0 + 1 & ", S: " & s1 & ", I: " & s2 + 1 & ", O:" & strT2 & "} " & strT(1) & " wait: " & strT(0), True
End If
GoSub TrO
SubmitIW Replace(strT(0), ",", ".") & "," & s0 & "," & s1 & "," & s2 & "," & s3 & "," & s4 & "," & ns(1) & "," & strT1 & "," & strT3 & "," & CInt(bolA) & "," & CInt(bolEn) & "," & strS
Exit Sub
End If
End If
Else
If bolDebug Then strT(0) = "T: " & s0 + 1 & ", S: " & s1 & ", I:" & s2 + 1 & ", O:" & strT2
Dim bytT As Byte
If UBound(lngProxyPos) = 0 Then bytT = 0 Else: bytT = s0
If lngProxyPos(bytT) > -1 Then
If SetProxy(s4, CByte(s0), strT(0)) Then
lngProxyPos(bytT) = -1
If strWait(ns(0), s2) <> vbNullString Then
If frmPT.bolSame Then addLog "{T: " & s0 & "} Warning: Out of working/valid proxies!" Else: addLog "Warning: Out of working/valid proxies!"
GoTo C
Else
If frmPT.bolSame Then addLog "{T: " & s0 & "} Error: Out of working/valid proxies!" Else: addLog "Error: Out of working/valid proxies!"
GoTo E
End If
End If
Else: If strWait(ns(0), s2) <> vbNullString Then GoTo C Else: GoTo E
End If
End If
GoSub TrO
Finish s0, s1, s2, s3, s4, ns(1), strS, strT1, strT3, bolA, bolEn
Exit Sub
TrO:
PrepTr s0, s1, s2, s3, strTNum, bolEn, strS, N, strT3
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
If frmT.bytTOrigin0 > 0 Then
Dim lngC As Long: lngC = Len(strT1) - Len(Replace(strT1, " ", vbNullString)) - frmT.bytTOrigin0
If lngC > 0 Then
strT1 = Replace(strT1, " ", vbNullString, , lngC)
strT1 = " ..." & Mid$(strT1, InStr(strT1, " "))
End If
End If
Return
End Sub

Private Sub PrepTr(s0 As String, s1 As String, s2 As String, s3 As String, strTNum As String, Optional bolEn As Boolean, Optional strS As String, Optional N As Integer, Optional strT3 As String)
Dim a As Integer, strT1 As String
Static strT As String
strT = strT & s0 & "," & strTNum & "\" & vbLf
If Len(strT) - Len(Replace(strT, vbLf, vbNullString)) = 1 Then
Do While strT <> vbNullString
strT1 = Left$(strT, InStr(strT, vbLf) - 1)
For a = colStr.count To 1 Step -1
If InStr(colStr.Item(a), strT1) = 1 Then colStr.Remove a
Next
strT = Mid$(strT, Len(strT1) + 2)
Loop
End If
If bytOrigin > 0 Then strCurrO = Replace(strCurrO, "," & s0 & "," & s3 & "," & vbLf, vbNullString, , 1)
strT3 = TrimO(s3 & " " & s1, strTNum, s0, N)
If bolEn Or strS = vbNullString Then Exit Sub
If bytOrigin = 0 Then
On Error Resume Next
colSrc.Remove s0 & "," & s2 & "," & strT3
On Error GoTo 0
colSrc.add strS, s0 & "," & s2 & "," & strT3
Else
On Error Resume Next
colSrc.Remove s0 & "," & strT3
On Error GoTo 0
colSrc.add s0 & "," & strT3 & "\" & strS, s0 & "," & strT3
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
obj(i).Tag = strT
intTmrCount = intTmrCount + 1
obj(i).Enabled = True
Set obj = Nothing
Exit Sub
E:
i = i + 1
Resume C
End Sub

Private Sub Finish(s0 As String, s1 As String, s2 As String, s3 As String, s4 As String, ns As Byte, strS As String, strT1 As String, strT As String, bolA As Boolean, bolEn As Boolean)
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
If bytOrigin > 0 Then colMax.Remove s2 & "," & s0 & "," & s3
End If
End If
If bytOrigin > 0 Then If InStr(strCurrO, "," & s0 & "," & strT & ",") = 0 Then strCurrO = "," & s0 & "," & strT & "," & vbLf
Dim a As Integer
For a = 0 To Val(strM)
If bolAb Or ChkSt Then Exit Sub
If strURLData(ns) = vbNullString Then
SubmitIW s0 & "," & a & "," & ns & "," & strT & "," & strT1 & "," & TrimO(strT & " " & a) & "," & strS, 1
'Else: SubmitIW s0 & ",0," & ns & "," & s3 & " " & s1 & "," & strT1 & "," & strS, 1
'End If
Else: SubmitReq ns, CByte(s0), s4, strT, CInt(a), strT1
End If
Next
End If
End Sub

Private Function TrimO(strT As String, Optional strTNum As String, Optional j As String, Optional N As Integer) As String 'to be fixed
Dim b As Byte
'del
'Open "problem.txt" For Append As #1
'Print #1, "BEFORE (" & colSrc.count & "):"
'For b = 1 To colSrc.count
'Debug.Print colSrc.Item(b) 'Print #1, Left$(colSrc.Item(b), InStr(colSrc.Item(b), "\") - 1)
'Next
'Debug.Print
'b = 0
'del
If bytOrigin > 0 Then
'Debug.Print lngC, bytOrigin
If Len(strT) - Len(Replace(strT, " ", vbNullString)) - bytOrigin = bytOrigin Then
Static bytT(1) As Byte, colR(1) As Collection, strC(1) As String
If j <> vbNullString Then
bytT(0) = bytT(0) + 1
If bytT(0) = 1 Then Set colR(0) = New Collection
End If
TrimO = strT
If Left$(strT, 1) = "/" Then strT = "/" Else: strT = vbNullString
ShortO TrimO, bytOrigin, strT
If j <> vbNullString Then
Dim a As Integer, i As Integer, colT As Collection, varT As Variant, strT1(6) As String, intC(1) As Integer, strJ As String, bytJ As Byte
strJ = j & ","
bytJ = Len(strJ)
Set colT = New Collection
Dim obj As Collection
Set obj = colPubStr
R:
For i = obj.count To 1 Step -1
strT1(0) = obj.Item(i)
PullK b = 0, strT1(0), strT1(6), a, strT1(4)
If StrPtr(PrepareCol(colR(b), strT1(6))) = 0 Then
Select Case True
Case b = 1, Compare(strTNum, strT1(4)) <> -1
If a > 1 Then strT1(1) = Mid$(strT1(6), InStr(strT1(6), "%") + 1) Else: strT1(1) = strT1(6)
If Left$(strT1(1), InStr(strT1(1), ",") - 1) = j Then
strT1(1) = Left$(strT1(0), InStr(a, strT1(0), ","))
strT1(2) = Mid$(strT1(0), Len(strT1(1)) + 1)
If Left$(strT1(2), 1) = "/" Then strT1(5) = "/" Else: strT1(5) = vbNullString
strT1(3) = Left$(strT1(2), InStr(strT1(2), "\") - 1)
If InStr(strCurrO, "," & j & "," & strT1(3)) = 0 Then
If InStr(tmrQ.Tag, "," & j & "," & strT1(3)) = 0 Then
strT1(4) = strT1(3)
Do
intC(0) = InStrRev(strT1(3), " ", Len(strT1(4)) - 1)
intC(1) = InStrRev(strT1(3), "|", Len(strT1(4)) - 1)
If intC(0) = 0 And intC(1) = 0 Then Exit Do
If intC(0) > intC(1) Then strT1(4) = Left$(strT1(3), intC(0) - 1) Else: strT1(4) = Left$(strT1(3), intC(1) - 1)
If Len(strT1(4)) <= 1 Then Exit Do
Select Case True
Case InStr(strCurrO, "," & j & "," & strT1(4)) > 0, InStr(tmrQ.Tag, "," & j & "," & strT1(4)) > 0: GoTo N
End Select
Loop
If strT1(3) <> strT1(5) & "..." Then
strT1(4) = strT1(3)
ShortO strT1(4), bytOrigin, strT1(5)
'Else
'lngC = InStr(strT1(0), strT1(3))
'If lngC <> 1 And lngC <> 2 Then
'lngC = CInt(Mid$(strT1(0), Len(strT1(0)) - Len(strT1(3)) + Len(strT1(2)) + 4) = Mid$(strT1(3), Len(strT1(2)) + 4))
'If lngC = -1 Then strT1(4) = strT1(5) & "..." Else: GoTo N
'Else: strT1(4) = strT1(5) & "..."
'End If
'End If
strT1(0) = Left$(strT1(0), Len(strT1(1))) & strT1(4)
If strT1(4) = strT1(5) & "..." Then
If b = 0 Then strT1(1) = Left$(strT1(1), InStr(strT1(1), "%"))
For Each varT In colT
If InStr(varT, strT1(1)) = 1 Then
If b = 0 Then
If Mid(varT, InStr(varT, "{") + 1, bytJ) = strJ Then GoTo N1
Else: GoTo N1
End If
End If
Next
strC(b) = strC(b) & strT1(1) & vbLf
ElseIf PrepareCol(colT, strT1(0)) <> vbNullString Then GoTo N
End If
'Debug.Print "+"; Left$(strT1(0) & vbNewLine, InStr(strT1(0) & vbNewLine, vbNewLine) - 1)
colT.add strT1(0) & Mid$(strT1(2), Len(strT1(3)) + 1), strT1(0)
N1:
On Error Resume Next
colR(b).add strT1(6), strT1(6)
On Error GoTo 0
Else
If b = 0 Then strT1(1) = Left$(strT1(1), InStr(strT1(1), "%")) & strJ
If InStr(vbLf & strC(b), vbLf & strT1(1) & vbLf) > 0 Then GoTo N1 Else: strC(b) = strC(b) & strT1(1) & vbLf ': Debug.Print "R: "; Left$(strT1(1) & vbNewLine, InStr(strT1(1) & vbNewLine, vbNewLine) - 1)
': Debug.Print "S: "; Left$(strT1(1) & vbNewLine, InStr(strT1(1) & vbNewLine, vbNewLine) - 1)
End If
End If
End If
N:
End If
End Select
End If
Next
On Error Resume Next
For i = colT.count To 1 Step -1
strT1(0) = colT.Item(i)
PullK b = 0, strT1(0), strT1(1)
obj.Remove strT1(1)
obj.add strT1(0), strT1(1) ', obj.count ', 1
colR(b).Remove strT1(1)
Next
On Error GoTo 0
Set colT = Nothing
bytT(b) = bytT(b) - 1
i = 1
Do While bytT(b) = 0 And colR(b).count >= i
'Debug.Print "-"; colR(b).Item(i)
obj.Remove colR(b).Item(i)
i = i + 1
Loop
If bytT(b) = 0 Then Set colR(b) = Nothing: strC(b) = vbNullString
'For i = 1 To obj.count
'Debug.Print Left$(obj.Item(i) & vbNewLine, InStr(obj.Item(i) & vbNewLine, vbNewLine) - 1)
'Next
'Debug.Print IIf(obj Is colSrc, "SRC: ", "PUBSTR: "); obj.count & vbNewLine
If b = 0 Then
Set obj = colSrc
Set colT = New Collection
If bytT(1) = 0 Then Set colR(1) = New Collection
b = 1
a = 1
bytT(1) = bytT(1) + 1
GoTo R
End If
'Debug.Print colPubStr.count, colSrc.count & vbNewLine
End If
GoTo E
End If
End If
If j <> vbNullString Then If Left$(strT, 1) <> "/" Then If Mid$(strT, InStrRev(strT, " ") + 1, 1) <> "0" Then TrimO = "/"
TrimO = TrimO & strT
E:
If N < 0 Then
If frmT.bolColl Then
Dim lngC As Long
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
'del
'Debug.Print strT, " -> ", TrimO
'If b = 0 Then Exit Function
'Open "problem.txt" For Append As #1
'Print #1, "AFTER (" & colPubStr.count & "):"
'For i = 1 To colPubStr.count
'Print #1, colPubStr.Item(i) 'Left$(colPubStr.Item(i), InStr(colPubStr.Item(i), "\") - 1)
'Next
'Print #1,
'Close #1
Debug.Print colPubStr.count, colSrc.count
End Function

Private Sub ShortO(strIO As String, bytS As Byte, strT As String)
Dim intC(2) As Integer, i As Integer
intC(0) = 1
intC(1) = 1
intC(2) = 1
Do
If intC(0) > 0 Then intC(0) = InStr(intC(2), strIO, " ")
If intC(1) > 0 Then intC(1) = InStr(intC(2), strIO, "|")
If intC(0) = 0 And intC(1) = 0 Then Exit Do
If (intC(1) > intC(0) Or intC(1) = 0) And intC(0) > 0 Then intC(2) = intC(0) + 1 Else: intC(2) = intC(1) + 1
If i = bytS Then Exit Do
i = i + 1
Loop
If i = bytS Then strIO = strT & "..." & Mid$(strIO, intC(2) - 1) Else: strIO = strT & "..."
End Sub

Private Sub PullK(bolT As Boolean, strI As String, strO As String, Optional a As Integer, Optional strTNum As String)
Dim b(1) As Integer
If bolT Then
strO = Left$(strI, InStr(strI, "%"))
b(0) = Len(strO) + 1
b(1) = InStr(b(0), strI, ",")
strTNum = Mid$(strI, b(0), b(1) - b(0))
a = InStr(b(1), strI, "{") + 1
strO = strO & Mid$(strI, a, InStr(a, strI, ",") - a) & "," & strTNum
Else: strO = Left$(strI, InStr(strI, "\") - 1)
End If
End Sub

Private Sub tmrQ_Timer()
Dim s(7) As String, i As Byte, ln(1) As Long, strT As String
If Not bolAb And tmrQ.Tag <> vbNullString Then
If 255 - bytActive < 1 Then Exit Sub
strT = Mid$(tmrQ.Tag, 2, InStr(tmrQ.Tag, vbLf) - 2)
ln(0) = 1
For i = 0 To 7
ln(1) = InStr(ln(0), strT, ",")
If ln(1) = 0 Then ln(1) = Len(strT) + 1
s(i) = Mid$(strT, ln(0), ln(1) - ln(0))
ln(0) = ln(1) + 1
Next
Dim intT As Integer
If s(6) = "0" And Left$(s(2), 1) <> "/" Then intT = 255 Else: intT = intSubT
If intT - bytActive > 0 Then SubmitReq CByte(s(0)), CByte(s(1)), s(5), s(2), CInt(s(6)), s(7), s(3), s(4) Else: Exit Sub
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
Dim s(10) As String, i As Byte, ln(1) As Long
ln(0) = InStr(tmrW(Index).Tag, ",") + 1
For i = 0 To 9
ln(1) = InStr(ln(0), tmrW(Index).Tag, ",")
s(i) = Mid$(tmrW(Index).Tag, ln(0), ln(1) - ln(0))
ln(0) = ln(1) + 1
Next
s(10) = Mid$(tmrW(Index).Tag, ln(0))
If s(10) <> vbNullChar Then
intTmrCount = intTmrCount - 1
Finish s(0), s(1), s(2), s(3), s(4), CByte(s(5)), s(10), s(6), s(7), CBool(s(8)), CBool(s(9))
If bolUnl Then Exit Sub
Else
intTmrCount = intTmrCount - 1
SubmitReq CByte(s(0)), CByte(s(1)), s(2), s(3), CInt(s(4)), s(5), s(6)
End If
Else
intTmrCount = intTmrCount - 1
If Not bolUnl Then tmrW(Index).Enabled = False
End If
If Not bolUnl Then If Index > 0 Then Unload tmrW(Index): intLTmr(1) = Index
If Not bolAb Then If s(10) <> vbNullChar Then If Not tmrQ.Enabled Then ChkT CBool(s(8)) Else: ChkSt
End Sub

Private Sub tmrU_Timer(Index As Integer)
tmrU(Index).Enabled = False
If Not bolAb Then
Dim strP As String
If Not frmPT.bolNoStartP Then
If frmPT.strProxy <> vbNullString Then
If SetProxy(strP, CByte(Index), "T: " & Index + 1 & ", S: 0, I: 1") Then
If Not bolAb Then
addLog "Error: Out of valid proxies!"
Enb 2
End If
intTmrCount = intTmrCount - 1
GoTo E
End If
End If
End If
intTmrCount = intTmrCount - 1
SubmitReq 0, CByte(Index), strP
Else: intTmrCount = intTmrCount - 1
End If
E: If Index > 0 Then Unload tmrU(Index)
End Sub

Private Sub tmrI_Timer(Index As Integer)
tmrI(Index).Enabled = False
If Not bolAb Then
Dim s(8) As String
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
If Not frmPT.bolNoStartP Then
If frmPT.strProxy <> vbNullString Then
If SetProxy(s(7), CByte(Index), "T: " & Index + 1 & ", S: 0, I: 1") Then
If Not bolAb Then
addLog "Error: Out of valid proxies!"
Enb 2
End If
intTmrCount = intTmrCount - 1
If Index > 0 Then Unload tmrI(Index)
Exit Sub
End If
End If
End If
End If
Dim bolM As Boolean: bolM = True
intTmrCount = intTmrCount - 1
StrAdd strNum(s(0)), "1"
s(8) = strNum(s(0))
CheckIf s(0), s(1), s(2), s(3), s(7), s(6), s(4), s(5), s(8), bolM
Else: intTmrCount = intTmrCount - 1
End If
If bolUnl Then Exit Sub
If Index > 0 Then
Unload tmrI(Index)
intLTmr(0) = Index 'If intLTmr = 0 Or index < intLTmr Then
'Else: tmrI(Index).Tag = vbNullString
End If
If Not tmrQ.Enabled Then ChkT bolM Else: ChkSt 'If Not bolAb Then
End Sub

Private Function PrepareCol(col As Collection, k As String) As String
On Error Resume Next
PrepareCol = col.Item(k)
End Function

Private Function ChkStr(strInp As String, Optional bytP As Byte, Optional a As Byte, Optional bytY As Byte, Optional i As Byte, Optional bolT As Boolean) As Boolean
Dim intC(2) As Long
If Left$(strInp, 1) = "[" Then
ChkStr = True
Exit Function
ElseIf Left$(strInp, 1) = "'" Then
intC(0) = FindC1(strInp)
If intC(0) <> Len(strInp) Then If Mid$(strInp, intC(0) + 1, 1) <> "+" Then ChkStr = True
If bytP = 0 Then Exit Function Else: GoTo C
Else: If Left$(strInp, 1) = "%" Then If Mid$(strInp, 2, 1) <> "{" Then If InStr(2, strInp, "%") = Len(strInp) Then Exit Function
End If
If InStr("%<", Left$(strInp, 1)) = 0 Then
If InStr(strInp, "(") > 0 Then
If InStr(Comms & frmPlugins.strC, "," & Left$(strInp, InStr(strInp, "(") - 1) & ",") = 0 Then ChkStr = True
Else: ChkStr = True
End If
If bytP <> 0 Then GoTo C
ElseIf Left$(strInp, 1) = "<" And bytP = 0 Then
ChkStr = Mid$(strInp & "+", InStr(2, strInp, ">") + 1, 1) <> "+"
If bytP = 0 Then Exit Function Else: GoTo C
Else
C:
Dim bytT As Byte, strT As String
intC(0) = 1
intC(2) = 1
Do
FindStr strInp, intC
If intC(1) <= 0 Or intC(0) > intC(1) Then Exit Do
If Mid$(strInp, intC(1), 1) = "}" Then
If bytP = 2 Then ChkStr Mid$(strInp, intC(0), intC(1) - intC(0)), 2, a, bytY, i, bolT Else: If bytP = 1 And ChkStr Then If Not ChkStr(Mid$(strInp, intC(0), intC(1) - intC(0)), 1) Then ChkStr = False: Exit Function
bytT = 2
Else: bytT = 1
End If
If Not ChkStr Then
If intC(2) > 1 Then strT = Mid$(strInp, intC(2), 1) Else: strT = vbNullString
strT = Replace(strT & Mid$(strInp, intC(1) + bytT, 1), "+", vbNullString)
If intC(0) > bytT + 1 Then If InStr("(,", Mid$(strInp, intC(0) - bytT - 1, 1)) > 0 Then strT = Replace(Replace(strT, ",", vbNullString), ")", vbNullString, , 1)
ChkStr = strT <> vbNullString
End If
If bytP <> 0 Then
If bytP = 2 Then strInp = Left$(strInp, intC(0) - bytT - 1) & "''" & Mid$(strInp, intC(1) + bytT)
If intC(1) = Len(strInp) Then Exit Do
If bytP <> 2 Then intC(0) = intC(1) + 1 Else: intC(0) = intC(1) - (intC(1) - intC(0)) + 1
intC(2) = intC(0)
Do While Mid$(strInp, intC(2), 1) = ")" And intC(2) < Len(strInp)
intC(2) = intC(2) + 1
Loop
If intC(2) = Len(strInp) Then Exit Do
Else: Exit Function
End If
Loop
If bytP = 2 And Not ChkStr Then If bolDebug Then PlugCheck strInp, a, bytY, i, bolT Else: PlugCheck strInp
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

Private Sub Form_KeyPress(KeyAscii As Integer)
If ActiveControl Is Nothing Then Exit Sub
Const ASC_CTRL_A As Integer = 1

    ' See if this is Ctrl-A.
    Select Case KeyAscii
      Case ASC_CTRL_A
      KeyAscii = 0
        ' The user is pressing Ctrl-A. See if the
        ' active control is a TextBox.
        If TypeOf ActiveControl Is TextBox Then
            ' Select the text in this control.
            ActiveControl.SelStart = 0
            ActiveControl.SelLength = Len(ActiveControl.Text)
        End If
      Case vbKeyEscape
        If Not TypeOf ActiveControl Is TextBox And ActiveControl.Name <> "cmbField" Then Exit Sub
        Dim strT As String
        Select Case ActiveControl.Name
            Case "txtName": strT = strName(bytI)
            Case "txtURL": strT = Split(strURLData(bytI) & vbLf, vbLf)(0)
            Case "txtData": strT = Split(strURLData(bytI) & vbLf, vbLf)(1)
            Case "txtA": strT = Split(strIf(bytI, ActiveControl.Index) & vbLf, vbLf)(0)
            Case "txtB": strT = Split(strIf(bytI, ActiveControl.Index) & vbLf & vbLf, vbLf)(2)
            Case "txtString": strT = Split(strStrings(bytI, ActiveControl.Index) & vbLf, vbLf)(0)
            Case "txtExp": strT = Split(strStrings(bytI, ActiveControl.Index) & vbLf, vbLf)(1)
            Case "cmbField": strT = Split(strHeaders(bytI, ActiveControl.Index) & vbLf, vbLf)(0)
            Case "txtValue": strT = Split(strHeaders(bytI, ActiveControl.Index) & vbLf, vbLf)(1)
            Case "txtWait": strT = strWait(ActiveControl.Index, bytI)
        End Select
        If ActiveControl.Text = strT Then Exit Sub
        ActiveControl.Text = strT
        ActiveControl.SelStart = 0
        ActiveControl.SelLength = Len(ActiveControl.Text)
      Case vbKeyReturn
        If Not TypeOf ActiveControl Is TextBox And TypeOf ActiveControl Is ComboBox And ActiveControl.Name <> "cmbField" And ActiveControl.Name <> "txtExp" Then Exit Sub
        Select Case ActiveControl.Name
            Case "txtName": txtName_Validate False
            Case "txtURL": txtURL_Validate False
            Case "txtData": txtData_Validate False
            Case "txtA": txtA_Validate ActiveControl.Index, False
            Case "txtB": txtB_Validate ActiveControl.Index, False
            Case "txtString": txtString_Validate ActiveControl.Index, False
            'Case "txtExp": txtExp_Validate ActiveControl.Index, False
            'Case "cmbField": cmbField_Validate ActiveControl.Index, False
            Case "txtValue": txtValue_Validate ActiveControl.Index, False
            Case "txtWait": txtWait_Validate ActiveControl.Index, False
        End Select
    End Select
End Sub

Private Sub chkOnTop_Click()
Static tm As Boolean
tm = Not tm
  If tm Then
  chkOnTop.Checked = True
    SetTopMostWindow Me.hWnd, True
    Me.Caption = Me.Caption & " [on top]"
  Else
  chkOnTop.Checked = False
    SetTopMostWindow Me.hWnd, False
    Me.Caption = Left$(Me.Caption, InStrRev(Me.Caption, " [on top]") - 1) & Mid$(Me.Caption, InStrRev(Me.Caption, " [on top]") + 9)
  End If
End Sub

Private Sub cmdNew_Click()
If Left$(Me.Caption, 1) = "*" And Filled(0) Then
Select Case MsgBox("Save current configuration?", vbYesNoCancel + vbExclamation)
Case vbYes: cmdSaveC_Click
Case vbCancel: Exit Sub
End Select
End If
RemC
End Sub

Private Sub RemC(Optional bolS As Boolean)
cmdOpt(0).Tag = vbNullString
cmdSaveC.Tag = vbNullString
lblStatus.Caption = "Removing current configuration..."
lblStatus.Refresh
'disF
If Not bolS Then Screen.MousePointer = 11
cmbIndex.Clear
cmbIndex.AddItem "1"
Dim i As Byte
For i = 0 To 1
cmbGoto(i).Clear
cmbGoto(i).AddItem "Next"
cmbGoto(i).AddItem "1"
Next
DimP True
bolChk = False
bolL = False
strPl = vbLf
cmbIndex.ListIndex = 0
cmdNew.Enabled = False
cmbIndex.Enabled = True
EnbC True
cmdR.Enabled = False
Me.Caption = "UniBot"
cmdStart.Tag = vbNullString
cmdStart.Enabled = False
If bolDebug Then addLog "Current configuration removed.", True
If bolS Then Exit Sub
Screen.MousePointer = 0
lblStatus.Caption = "Idle..."
End Sub

Private Sub chkProxy_Click(Index As Integer)
If Left$(lblStatus.Caption, 5) = "Shift" Or Left$(lblStatus.Caption, 4) = "Load" Or Left$(lblStatus.Caption, 5) = "Remov" Then Exit Sub
bolProxy(Index, bytI) = chkProxy(Index).Value
RplTitle
If bolDebug Then addLog "Proxy {index: " & bytI + 1 & ", number: " & Index + 1 & "}: " & chkProxy(Index).Value, True
End Sub

Private Sub cmdMintoTray_Click()
Me.Visible = False
If App.LogMode > 0 Then
SystemTray.Tip = Me.Caption
SystemTray.AddIcon
End If
cmdMintoTray.Checked = True
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
strT(1) = strT(1) & "|HTML file (*.html)|*.html"
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
strT(1) = vbNullString 'strLog
Dim i As Integer
For i = 0 To lstLog.ListCount - 1
strT(1) = strT(1) & lstLog.list(i) & vbCrLf
Next
strT(1) = Left$(strT(1), Len(strT(1)) - 2)
Print #1, strT(1);
Close #1
Else: PutContents strFile, txtOutput.Text, IIf(IsUnicode(txtOutput.Text), CP_UTF16_LE, CP_ACP)
End If
If bolDebug Then addLog UCase$(Left$(strT(2), 1)) & Mid$(strT(2), 2) & " saved (" & get_relative_path_to(strFile) & ").", True
lblStatus.Caption = "Idle..."
Exit Sub
E:
Close #1
If bolDebug Then addLog "Failed to save " & strT(2) & ".", True
lblStatus.Caption = "Error! Idle..."
If strPath(Index) = vbNullString Then MsgBox "Error in saving " & strT(2) & " file!", vbCritical
End Sub

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

Sub PutContents(ByRef the_sFilename As String, ByRef the_sFileContents As String, ByVal the_nCodePage As Long, Optional the_bPrefixWithBOM As Boolean = True)

    Dim iFileNo                     As Integer

    iFileNo = FreeFile
    OpenForOutput the_sFilename, iFileNo, the_nCodePage, the_bPrefixWithBOM
    Print_ iFileNo, the_sFileContents, the_nCodePage, False
    Close iFileNo

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
If Not Me.Visible And App.LogMode > 0 Then SystemTray.RemoveIcon
Set SystemTray = Nothing
Set Plugins = Nothing
Dim s() As String, i As Integer ', cnt As Integer
s() = Split(frmPlugins.strP, vbLf)
For i = 0 To UBound(s) - 1
UnloadLibrary Split(s(i), "|")(0) & Split(s(i), "|")(1)
Next
If frmPlugins.strRg <> vbNullString Then Close #2
If chkNoSave.Checked Then Exit Sub
If strLastPath <> vbNullChar Then If CurDir$ <> strInitD Then strLastPath = get_relative_path_to(CurDir$, True) Else: strLastPath = vbNullString
If bytLimit <> 19 Or frmPT.bytTimeout <> 20 Or frmPT.bytSubThr <> 1 Or frmPT.bolSkip Or frmPT.bolSame Or frmPT.bolNoStartP Or frmPT.bolNoRetry Or frmPT.bolNoChange Or frmPT.bytDelay <> 1 Or frmPT.bytMaxR > 0 Or frmPT.bytCycles > 0 Or frmPlugins.strRL <> vbNullString Or strLastPath <> vbNullChar Or frmT.strTemplate0 <> vbNullString Or frmT.strTemplate1 <> vbNullString Or bolDebug Or frmT.bytTOrigin0 <> 5 Or frmT.bytTOrigin1 <> 5 Or frmT.bytTOrigin1 > 0 And frmT.bolNoEach Or frmT.bolColl Or frmT.intLogMax < 32767 Or frmT.intOutMax > 0 Or frmWizard.strUA <> DEFUSERAGENT Or bolSkipLE = vbTrue Or bolSilent2 = vbTrue Then 'Or frmPT.bytThreads <> 1
On Error GoTo E
Open strInitD & "Settings.ini" For Output Access Write As #1
Print #1, BuildS;
E:
Close #1
If bolHid Then SetAttr strInitD & "Settings.ini", vbHidden
ElseIf Dir$(strInitD & "Settings.ini") <> vbNullString Then Kill strInitD & "Settings.ini"
End If
End Sub

Private Function BuildS(Optional intE As Integer = -1) As String
If frmPT.bytTimeout <> 20 Then BuildS = "timeout=" & frmPT.bytTimeout & vbNewLine
If frmPT.bytSubThr <> 1 Then BuildS = BuildS & "subthreads=" & frmPT.bytSubThr & vbNewLine
If frmPT.bolNoRetry Then BuildS = BuildS & "donotretry=1" & vbNewLine
If frmPT.bytDelay <> 1 Then BuildS = BuildS & "delaybetweenretries=" & frmPT.bytDelay & vbNewLine
If frmPT.bytMaxR > 0 Then BuildS = BuildS & "maxretriespercycle=" & frmPT.bytMaxR & vbNewLine
If frmPT.bytCycles > 0 Then BuildS = BuildS & "maxcycles=" & frmPT.bytCycles & vbNewLine
If bolDebug Then BuildS = BuildS & "debug=1" & vbNewLine
Dim strT As String
If frmT.strTemplate0 <> vbNullString Or frmT.strTemplate1 <> vbNullString Then
If frmT.strTemplate1 <> vbNullString Or strT <> vbNullString Then strT = Replace(Replace(frmT.strTemplate1, """", strC), vbNewLine, "[nl]") & ""","""
If frmT.strTemplate0 <> vbNullString Or strT <> vbNullString Then strT = Replace(Replace(frmT.strTemplate0, """", strC), vbNewLine, "[nl]") & """,""" & strT
BuildS = BuildS & "output=" & """" & Left$(strT, Len(strT) - 2) & vbNewLine
End If
If frmT.intAfter > 0 Then
BuildS = BuildS & "after=" & frmT.intAfter
If frmT.bolHours Then BuildS = BuildS & "h" & vbNewLine Else: BuildS = BuildS & vbNewLine
End If
If frmT.bytTOrigin0 <> 5 Or frmT.bytTOrigin1 <> 5 Or frmT.bolNoEach Or frmT.bolColl Then
BuildS = BuildS & "originmax="
If frmT.bytTOrigin0 <> 5 Then BuildS = BuildS & frmT.bytTOrigin0
If frmT.bytTOrigin1 <> 5 Then BuildS = BuildS & ";" & frmT.bytTOrigin1
If frmT.bolNoEach Then BuildS = BuildS & "n"
If frmT.bolColl And frmT.bytTOrigin1 > 0 Then BuildS = BuildS & "c"
BuildS = BuildS & vbNewLine
End If
If frmT.intLogMax <> 32767 Or frmT.intOutMax <> 0 Then
BuildS = BuildS & "logoutputmax="
If frmT.intLogMax <> 32767 Then BuildS = BuildS & frmT.intLogMax
If frmT.intOutMax <> 0 Then BuildS = BuildS & ";" & frmT.intOutMax
BuildS = BuildS & vbNewLine
End If
Dim s() As String, s1() As String, a As Byte, i As Integer
If intE = -1 Then
If bytLimit <> 19 Then BuildS = BuildS & "limit=" & bytLimit + 1 & vbNewLine
If frmPT.bolSame Then BuildS = BuildS & "sameproxyforeachthread=1" & vbNewLine
If frmPT.bolSkip Then BuildS = BuildS & "skipbadproxies=1" & vbNewLine
If frmPT.bolNoStartP Then BuildS = BuildS & "donotstartwithproxy=1" & vbNewLine
If frmPT.bolNoChange Then BuildS = BuildS & "donotchangeproxy=1" & vbNewLine
If strLastPath <> vbNullChar Then BuildS = BuildS & "lastpath=" & strLastPath & vbNewLine
If frmPlugins.strRL <> vbNullString Then BuildS = BuildS & "pluginsdir=" & frmPlugins.strRL & vbNewLine
If frmPlugins.strP <> vbNullString Then
s() = Split(frmPlugins.strP, vbLf)
For i = 0 To UBound(s) - 1
If InStr(s(i), "S|") > 0 Then
s1() = Split(s(i), "|")
BuildS = BuildS & "loadedplugin=" & get_relative_path_to(s1(0) & s1(1), , frmPlugins.strLocation) ' & cnt & "="
Dim b As Byte: If frmPlugins.strRg <> vbNullString Then b = 1
For a = 2 To UBound(s1) - 1
If Right$(s1(a), 1) = "S" Then BuildS = BuildS & "|" & Split(frmPlugins.strPl, vbLf)(Replace(s1(a), "S", vbNullString) - b)
Next
'cnt = cnt + 1
BuildS = BuildS & vbNewLine
End If
Next
End If
If frmWizard.strUA <> DEFUSERAGENT Then BuildS = BuildS & "defaultwizarduseragent=" & frmWizard.strUA & vbNewLine
If bolSkipLE = vbTrue Then BuildS = BuildS & "exeskiploadingerrors=2" & vbNewLine
If bolSilent2 = vbTrue Then BuildS = BuildS & "exemoresilent=2" & vbNewLine
Else
BuildS = BuildS & "limit=" & intE & vbNewLine
If frmPT.bytThreads <> 1 Then BuildS = BuildS & "threads=" & frmPT.bytThreads & vbNewLine
If strCmd <> vbNullString Then BuildS = BuildS & "executebatch=" & strCmd & vbNewLine
If strPath(0) <> vbNullString Then BuildS = BuildS & "savelog=" & get_relative_path_to(strPath(0), , CurDir$) & vbNewLine
If strPath(1) <> vbNullString Then BuildS = BuildS & "saveoutput=" & get_relative_path_to(strPath(1), , CurDir$) & vbNewLine
BuildS = BuildS & "title=" & frmEXE.txtTitle.Text & vbNewLine
If bolSkipLE <> vbFalse Then BuildS = BuildS & "skiploadingerrors=1" & vbNewLine
If frmEXE.optSilent Then
If bolSilent2 <> vbFalse Then BuildS = BuildS & "silent=2" & vbNewLine Else: BuildS = BuildS & "silent=1" & vbNewLine
ElseIf chkOnTop.Checked Then BuildS = BuildS & "ontop=1" & vbNewLine
End If
If frmEXE.chkMT Then BuildS = BuildS & "meltortray=1" & vbNewLine
If frmEXE.frm1.Enabled And frmEXE.lstPlugins.ListCount > 0 Then
If frmPlugins.strRg <> vbNullString And frmEXE.lstPlugins.Selected(0) Then
BuildS = BuildS & "loadedplugin=" & frmPlugins.strRg & vbNewLine
strPlg = App.path & IIf(Right$(App.path, 1) <> "\", "\", vbNullString) & frmPlugins.strRg & vbLf
End If
s() = Split(frmPlugins.strP, vbLf)
Dim strP(1) As String, strT1 As String, strL As String, strE As String
For i = 0 To UBound(s) - 1
s1() = Split(s(i), "|")
For a = 2 To UBound(s1) - 1
strT = Replace(s1(a), "S", vbNullString)
If frmEXE.lstPlugins.Selected(strT) Then
strT1 = vbNullString
Plugins.Item(Split(s(i), "|")(1) & "/" & frmEXE.lstPlugins.list(strT)).BuildSettings strT1
If strT1 <> vbNullString Then strE = strE & "[" & Split(s(i), "|")(1) & "/" & frmEXE.lstPlugins.list(strT) & "]" & vbNewLine & strT1 & vbNewLine Else: strL = strL & "|" & frmEXE.lstPlugins.list(strT)
If strP(0) = vbNullString Then frmPlugins.ExtrF CInt(strT), strP(0), 2
End If
Next
'cnt = cnt + 1
If strP(0) <> vbNullString Then
If strL <> vbNullString Then
strPlg = strPlg & strP(0) & vbLf
BuildS = BuildS & "loadedplugin=" & s1(1) & strL & vbNewLine
strL = vbNullString
Else: strP(1) = strP(1) & strP(0) & vbLf
End If
strP(0) = vbNullString
End If
Next
strPlg = strPlg & strP(1)
If strPlg <> vbNullString And frmEXE.optTemp Then BuildS = BuildS & "pluginsdir=1" & vbNewLine
BuildS = BuildS & strE
End If
End If
End Function

Private Sub SystemTray_MouseUp(Button As Integer)
 SetForegroundWindow Me.hWnd
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

Function Filled(bytIndex As Byte, Optional bytN As Byte) As Boolean
If bytN <> 1 Then
If strURLData(bytIndex) <> vbNullString Then Filled = Not ChkURL(Split(strURLData(bytIndex), vbLf)(0))
If Filled Then Exit Function
End If
If bytN <> 2 Then
If InStr(strIf(bytIndex, 0), vbLf) > 0 Then
If Split(strIf(bytIndex, 0), vbLf)(0) <> vbNullString Or Split(strIf(bytIndex, 0), vbLf)(2) <> vbNullString Then Filled = True
End If
If Filled Then Exit Function
End If
If bytN <> 3 Then If InStr(cmdOpt(0).Tag, "-" & bytIndex & "-") > 0 Then Filled = True
End Function

Private Sub cmbIndex_Click()
'On Error Resume Next
If lblStatus.Caption = "Starting..." Then Exit Sub
If Left$(lblStatus.Caption, 5) <> "Remov" And Left$(lblStatus.Caption, 4) <> "Load" And Left$(lblStatus.Caption, 5) <> "Shift" Then If ValidAll Then Exit Sub
If bytI <> cmbIndex.ListIndex Then
'If Left$(lblStatus.Caption, 5) <> "Remov" Then disF
bytI = cmbIndex.ListIndex
ElseIf Left$(lblStatus.Caption, 5) <> "Shift" And Left$(lblStatus.Caption, 5) <> "Remov" And Left$(lblStatus.Caption, 4) <> "Load" Then Exit Sub
End If
If VScroll1(0).Enabled Then VScroll1(0).Value = 0
If VScroll1(1).Enabled Then VScroll1(1).Value = 0
If HScroll1.Enabled Then HScroll1.Value = 0
Dim i As Byte
If Left$(lblStatus.Caption, 5) = "Remov" Then
txtName.Text = vbNullString
txtURL.Text = vbNullString
txtURL.Tag = vbNullString
txtData.Text = vbNullString
fraH.Tag = vbLf
fraS.Tag = vbLf
For i = 0 To 1
txtWait(i).Text = vbNullString
chkProxy(i).Value = 0
Next
cmbThen(1).ListIndex = 0
cmbGoto(1).ListIndex = 1
cmbThen(0).ListIndex = 0
cmbGoto(0).ListIndex = 0
cmbField(0).Tag = vbLf & "Cookie" & vbLf & "User-Agent" & vbLf
End If
For i = 0 To cmbField.count - 1
cmbField(i).Text = vbNullString
cmbField(i).Clear
txtValue(i).Text = vbNullString
Next
VScroll1(0).Value = 0
cmdAdd(0).Enabled = False
For i = 0 To txtExp.count - 1
txtString(i).Text = vbNullString
txtString(i).Tag = vbNullString
txtExp(i).Text = vbNullString
txtExp(i).Tag = vbNullString
cmdOpt(i).Enabled = False
Next
VScroll1(1).Value = 0
cmdAdd(1).Enabled = False
bolAl = True
For i = 0 To txtA.count - 1
txtA(i).Text = vbNullString
cmbSign(i).ListIndex = 0
txtB(i).Text = vbNullString
If i > 0 Then cmbOper(i - 1).ListIndex = 0
Next
HScroll1.Value = 0
bolAl = False
cmdAdd(2).Enabled = False
If Left$(lblStatus.Caption, 5) = "Remov" Then
'disF
If txtName.Enabled Then txtName.SetFocus
Exit Sub
End If
If Left$(lblStatus.Caption, 5) <> "Shift" Then
lblStatus.Caption = "Loading index number " & bytI + 1 & "..."
lblStatus.Refresh
Screen.MousePointer = 11
End If
txtName.Text = strName(bytI)
cmbField(0).Tag = vbLf
If InStr(strURLData(bytI), vbLf) > 0 Then
txtURL.Text = Split(strURLData(bytI), vbLf)(0)
txtData.Text = Split(strURLData(bytI), vbLf)(1)
Else
txtURL.Text = strURLData(bytI)
txtData.Text = vbNullString
End If
txtURL.Tag = txtURL.Text
Dim s() As String
i = 0
Do While strHeaders(bytI, i) <> vbNullString
s() = Split(strHeaders(bytI, i), vbLf)
If i > cmbField.count - 1 Then cmdAdd_Click 0
cmbField(i).Text = s(0)
txtValue(i).Text = s(1)
i = i + 1
If UBound(strHeaders, 2) < i Then Exit Do
Loop
If cmbField(cmbField.count - 1).Text <> vbNullString Then cmdAdd(0).Enabled = True
If fraH.Tag <> vbLf Then
Dim s1() As String
s1() = Split(fraH.Tag, vbLf)
Dim a As Byte
For a = 1 To UBound(s1()) - 1
If a Mod 2 <> 0 Then
cmbField(0).Tag = cmbField(0).Tag & Trim$(s1(a)) & vbLf
End If
Next
AddF , 1
End If
i = 0
Do While strStrings(bytI, i) <> vbNullString
s() = Split(strStrings(bytI, i), vbLf)
If i > txtExp.count - 1 Then cmdAdd_Click 1
txtString(i).Text = s(0)
txtExp(i).Text = s(1)
txtExp(i).Tag = s(2)
cmdOpt(i).Enabled = True
i = i + 1
If UBound(strStrings, 2) < i Then Exit Do
Loop
If txtExp(txtExp.count - 1).Text <> vbNullString Then cmdAdd(1).Enabled = True
Dim bytS As Byte
For i = 0 To UBound(strIf, 2)
If strIf(bytI, i) <> vbNullString Then
Dim j As Byte
For j = 1 To bytS
If i - j > txtA.count - 1 Then cmdAdd_Click 2
cmbOper(i - j - 1).ListIndex = 1
Next
bytS = 0
If i > txtA.count - 1 Then cmdAdd_Click 2
s() = Split(strIf(bytI, i), vbLf)
If i > 0 Then cmbOper(i - 1).ListIndex = s(3)
txtA(i).Text = s(0)
cmbSign(i).ListIndex = s(1)
txtB(i).Text = s(2)
ElseIf i > bytSh(bytI) Then Exit For
Else: bytS = bytS + 1
End If
Next
If txtA(txtA.count - 1).Text <> vbNullString Or txtB(txtA.count - 1).Text <> vbNullString Then cmdAdd(2).Enabled = True
txtWait(0).Text = strWait(0, bytI)
txtWait(1).Text = strWait(1, bytI)
chkProxy(0).Value = CInt(bolProxy(0, bytI)) * (-1)
chkProxy(1).Value = CInt(bolProxy(1, bytI)) * (-1)
If intGoto(0, bytI) <> 1 Then
cmbThen(0).ListIndex = 0
If intGoto(0, bytI) > 0 Then
If intGoto(0, bytI) - 1 <= cmbIndex.ListCount Then cmbGoto(0).ListIndex = intGoto(0, bytI) - 1 Else: GoTo D
Else: cmbGoto(0).ListIndex = 0
End If
cmbGoto(0).Enabled = True
Else
D:
cmbThen(0).ListIndex = 1
cmbGoto(0).ListIndex = 0
cmbGoto(0).Enabled = False
End If
If intGoto(1, bytI) > 0 Then
cmbThen(1).ListIndex = 1
If intGoto(1, bytI) - 1 <= cmbIndex.ListCount Then cmbGoto(1).ListIndex = intGoto(1, bytI) - 1 Else: GoTo D1
cmbGoto(1).Enabled = True
Else
D1:
cmbThen(1).ListIndex = 0
cmbGoto(1).ListIndex = bytI + 1
cmbGoto(1).Enabled = False
End If
'disF
If Left$(lblStatus.Caption, 5) <> "Shift" Then EnbIn 'If bolDebug Then addLog "Index " & bytI + 1 & " loaded.", True
End Sub

Private Sub EnbIn()
Screen.MousePointer = 0
If cmbIndex.Enabled And Me.Visible Then cmbIndex.SetFocus
lblStatus.Caption = "Idle..."
End Sub

Private Sub disF()
If Not fraI.Enabled Then
cmbIndex.Enabled = True
cmdManager.Enabled = True
txtName.Enabled = True
fraR.Enabled = True
fraI.Enabled = True
fraT.Enabled = True
fraS.Enabled = True
fraE.Enabled = True
conf.Enabled = True
cmdAutoSave.Enabled = True
cmdTuning.Enabled = True
cmdPlugins.Enabled = True
cmdAbout.Enabled = True
cmdI.Enabled = True
cmdR.Enabled = True
If lstLog.list(0) <> vbNullString Then
cmdSave(0).Enabled = True
cmdClear(0).Enabled = True
End If
If txtOutput.Text <> vbNullString Then
cmdSave(1).Enabled = True
cmdClear(1).Enabled = True
End If
cmdProxy.Enabled = True
If cmdStart.Tag = "1" Then
cmdStart.Enabled = True
cmdStart.Tag = vbNullString
End If
If cmdStart.Enabled Then
cmdManager.Enabled = True
cmdMake.Enabled = True
cmdShortcut.Enabled = True
Else
cmdManager.Enabled = False
cmdMake.Enabled = False
cmdShortcut.Enabled = False
End If
Else
cmbIndex.Enabled = False
cmdManager.Enabled = False
txtName.Enabled = False
fraR.Enabled = False
fraI.Enabled = False
fraS.Enabled = False
fraT.Enabled = False
fraE.Enabled = False
conf.Enabled = False
cmdAutoSave.Enabled = False
cmdPlugins.Enabled = False
cmdTuning.Enabled = False
cmdAbout.Enabled = False
cmdI.Enabled = False
cmdR.Enabled = False
If lstLog.list(0) <> vbNullString Then
cmdSave(0).Enabled = False
cmdClear(0).Enabled = False
End If
If txtOutput.Text <> vbNullString Then
cmdSave(1).Enabled = False
cmdClear(1).Enabled = False
End If
cmdProxy.Enabled = False
If cmdStart.Enabled Then
cmdManager.Enabled = False
cmdMake.Enabled = False
cmdShortcut.Enabled = False
cmdStart.Enabled = False
cmdStart.Tag = 1
End If
End If
End Sub

Private Sub cmbSign_Click(Index As Integer)
If Not bolAl And lblStatus.Caption <> "Starting..." And Left$(lblStatus.Caption, 5) <> "Shift" And Left$(lblStatus.Caption, 4) <> "Load" And Left$(lblStatus.Caption, 5) <> "Remov" Then If txtA(Index).Text <> vbNullString And txtB(Index).Text <> vbNullString Then AddIf Index
End Sub

Private Sub cmbOper_Click(Index As Integer)
If Not bolAl And lblStatus.Caption <> "Starting..." And Left$(lblStatus.Caption, 5) <> "Shift" And Left$(lblStatus.Caption, 4) <> "Load" And Left$(lblStatus.Caption, 5) <> "Remov" Then If txtA(Index).Text <> vbNullString And txtB(Index).Text <> vbNullString Then AddIf Index + 1, True
End Sub

Private Sub cmdAdd_Click(Index As Integer)
Dim intNext As Integer
Dim lngTL As Long
Select Case Index
Case 0
intNext = lbl1.count
If intNext > (bytLimit + 1) \ 2 Then If Not AddL Then Exit Sub
lngTL = cmbField(intNext - 1).Top + cmbField(0).Height + 45
Load cmbField(intNext)
Load lbl1(intNext)
Load txtValue(intNext)
With cmbField(intNext)
.ToolTipText = vbNullString
.Text = vbNullString
.Top = lngTL
.Left = cmbField(0).Left
.TabIndex = 10
.Visible = True
If lblStatus.Caption = "Idle..." Then .SetFocus
End With
With lbl1(intNext)
.Top = lngTL
.Left = lbl1(0).Left
.TabIndex = 11
.Visible = True
End With
With txtValue(intNext)
.ToolTipText = vbNullString
.Text = vbNullString
.Top = lngTL
.Left = txtValue(0).Left
.TabIndex = 12
.Visible = True
End With
PicBox1(0).Height = cmbField(intNext).Top + cmbField(0).Height
AddF intNext, 2
Case 1
intNext = lbl2.count
If intNext > (bytLimit + 1) \ 2 Then If Not AddL Then Exit Sub
lngTL = txtExp(intNext - 1).Top + txtExp(0).Height + 75
Load txtString(intNext)
Load lbl2(intNext)
Load txtExp(intNext)
Load cmdOpt(intNext)
With txtString(intNext)
.ToolTipText = vbNullString
.Text = vbNullString
.Top = lngTL
.Left = txtString(0).Left
.TabIndex = 21
.Visible = True
If lblStatus.Caption = "Idle..." Then .SetFocus
End With
With lbl2(intNext)
.Top = lngTL
.Left = lbl2(0).Left
.TabIndex = 22
.Visible = True
End With
With txtExp(intNext)
.Tag = vbNullString
.ToolTipText = vbNullString
.Text = vbNullString
.Top = lngTL
.Left = txtExp(0).Left
.TabIndex = 23
.Visible = True
End With
With cmdOpt(intNext)
.Top = lngTL
.Left = cmdOpt(0).Left
.Enabled = False
.TabIndex = 24
.Visible = True
End With
PicBox1(1).Height = txtExp(intNext).Top + txtExp(0).Height
Case 2
intNext = txtA.count
If intNext > (bytLimit + 1) \ 2 Then If Not AddL Then Exit Sub
lngTL = PicBox2.Width
If intNext - 1 > 0 Then Load cmbOper(intNext - 1)
Load txtA(intNext)
Load cmbSign(intNext)
Load txtB(intNext)
bolAl = True
If intNext > 1 Then
With cmbOper(intNext - 1)
.AddItem "Or"
.AddItem "And"
.ListIndex = 0
.Left = lngTL + 5
lngTL = .Left + .Width
.TabIndex = 27
.Visible = True
End With
Else
txtA(1).Width = txtA(1).Width - cmbOper(0).Width / 2
txtB(1).Width = txtA(1).Width
End If
bolAl = False
With txtA(intNext)
.ToolTipText = vbNullString
.Text = vbNullString
.Width = txtA(1).Width
.Left = lngTL
lngTL = .Left + .Width
.TabIndex = 28
.Visible = True
If lblStatus.Caption = "Idle..." Then .SetFocus
End With
With cmbSign(intNext)
.AddItem "="
.AddItem "<>"
.AddItem ">"
.AddItem "<"
.AddItem ">="
.AddItem "<="
.ListIndex = 0
.Left = lngTL + 5
lngTL = .Left + .Width
.ToolTipText = vbNullString
.TabIndex = 29
.Visible = True
End With
With txtB(intNext)
.ToolTipText = vbNullString
.Text = vbNullString
.Width = txtB(1).Width
.Left = lngTL
lngTL = .Left + .Width
.TabIndex = 30
.Visible = True
End With
PicBox2.Width = txtB(intNext).Left + txtB(1).Width
End Select
If Index <> 2 Then
If PicBox1(Index).Height > PicBox12(Index).Height Then
With VScroll1(Index)
.Enabled = True
.SmallChange = .Max \ (IIf(Index = 1, cmdOpt.count, txtValue.count) - 3)
.LargeChange = .SmallChange * 3
If lblStatus.Caption = "Idle..." Then
.Value = .Max
VScroll1_Scroll Index
End If
End With
End If
ElseIf PicBox2.Width > PicBox21.Width Then
With HScroll1
.Enabled = True
.SmallChange = .Max \ cmbOper.count
.LargeChange = .SmallChange
If lblStatus.Caption = "Idle..." Then
.Value = .Max
HScroll1_Scroll
End If
End With
End If
cmdAdd(Index).Enabled = False
End Sub

Private Function AddL(Optional bolS As Boolean) As Boolean
Dim intL As Integer
If Not bolS Then
If MsgBox("Current limit is reached! Force moving limit?", vbExclamation + vbYesNo) = vbNo Then bolL = True: Exit Function
If bytLimit + 2 > 255 Then
If bytLimit = 255 Then
MsgBox "No hard feelings, but you are complete lunatic... And I am too, for coding this shit."
Exit Function
End If
intL = 256 - bytLimit
Else
Dim strI As String: strI = InputBox("Enter new index limit:", , bytLimit + 3)
If strI = vbNullString Then Exit Function
If CInt(strI) = 0 Then Exit Function
intL = strI
If intL < 0 Then intL = intL * (-1)
If intL < 8 Then Exit Function
If intL < bytLimit + 3 Then intL = bytLimit + 3
If intL > 256 Then intL = 256
End If
lblStatus.Caption = "Moving limit..."
lblStatus.Refresh
'disF
ElseIf bytLimit + 3 <= 256 Then intL = bytLimit + 3
Else: Exit Function
End If
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
If frmManager.bolL Then frmManager.Tag = "  "
If bolDebug Then addLog "Limit moved, from " & bytL + 1 & " to " & intL & " (added " & intL - (bytL + 1) & " free spaces).", True
If Not bolS Then
'disF
Screen.MousePointer = 0
lblStatus.Caption = "Idle..."
End If
AddL = True
End Function

Private Sub cmbThen_Click(Index As Integer)
If lblStatus.Caption = "Starting..." Or Left$(lblStatus.Caption, 5) = "Shift" Or Left$(lblStatus.Caption, 4) = "Load" Or Left$(lblStatus.Caption, 5) = "Remov" Then Exit Sub
If Index = 0 Then
If cmbThen(Index).ListIndex = 1 Then
If intGoto(0, bytI) = 1 Then Exit Sub
intGoto(0, bytI) = 1
cmbGoto(0).Enabled = False
Else
Dim i As Byte
If cmbGoto(Index).ListIndex > 0 Then i = cmbGoto(Index).ListIndex + 1 Else: i = 0
If intGoto(0, bytI) = i Then Exit Sub Else: intGoto(0, bytI) = i
cmbGoto(0).Enabled = True
End If
ElseIf cmbThen(Index).ListIndex = 1 Then
If intGoto(1, bytI) = cmbGoto(Index).ListIndex + 1 Then Exit Sub
intGoto(1, bytI) = cmbGoto(Index).ListIndex + 1
cmbGoto(1).Enabled = True
ElseIf intGoto(1, bytI) <> 0 Then
If intGoto(1, bytI) = 0 Then Exit Sub
intGoto(1, bytI) = 0
cmbGoto(1).Enabled = False
Else: Exit Sub
End If
RplTitle
If bolDebug Then addLog "Go to {index: " & bytI + 1 & ", number " & Index + 1 & "}: " & intGoto(Index, bytI), True
End Sub

Private Sub cmdClear_Click(Index As Integer)
If MsgBox("Sure?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
If Index = 0 Then
'strLog = vbNullString
lstLog.Clear
SetListboxScrollbar
Else: txtOutput.Text = vbNullString
End If
cmdClear(Index).Enabled = False
cmdSave(Index).Enabled = False
End Sub

Private Sub cmdOpt_Click(Index As Integer)
frmD.Caption = txtString(Index).Text
frmD.Tag = Index
'Dim bolT As Boolean
If InStr(txtExp(Index).Tag, ",") > 0 Then
If Split(txtExp(Index).Tag, ",")(0) = "1" Then frmD.chkCrucial.Value = 1
If Split(txtExp(Index).Tag, ",")(1) = "1" Then frmD.chkPublic.Value = 1
If Split(txtExp(Index).Tag, ",")(2) = "1" Then frmD.chkArray.Value = 1
If Split(txtExp(Index).Tag, ",")(3) = "1" Then frmD.chkOutput.Value = 1
End If
If txtExp(Index).Tag = vbNullString Then txtString(Index).Tag = " "
frmD.Show vbModal
If txtExp(Index).Tag = vbNullString Then
If txtString(Index).Tag = vbNullString Then
txtString(Index).Tag = " "
'bolT = Split(Split(strStrings(bytI, index), vbLf)(2), ",")(0) = "1"
strStrings(bytI, Index) = Split(strStrings(bytI, Index), vbLf)(0) & vbLf & Split(strStrings(bytI, Index), vbLf)(1) & vbLf
RplTitle 'bolT
If bolDebug Then addLog "String {index: " & bytI + 1 & ", number: " & Index + 1 & "} -> (none)", True
End If
Exit Sub
ElseIf Left$(txtExp(Index).Tag, 1) = "-" Then
txtExp(Index).Tag = Mid$(txtExp(Index).Tag, 2)
Exit Sub
End If
Dim strT As String: strT = Split(strStrings(bytI, Index), vbLf)(2)
If strT <> vbNullString Then If Split(strT, ",")(0) = Split(txtExp(Index).Tag, ",")(0) And Split(strT, ",")(1) = Split(txtExp(Index).Tag, ",")(1) And Split(strT, ",")(2) = Split(txtExp(Index).Tag, ",")(2) And Split(strT, ",")(3) = Split(txtExp(Index).Tag, ",")(3) Then Exit Sub
strStrings(bytI, Index) = Split(strStrings(bytI, Index), vbLf)(0) & vbLf & Split(strStrings(bytI, Index), vbLf)(1) & vbLf & txtExp(Index).Tag
If Not cmbIndex.Enabled Then If Split(txtExp(Index).Tag, ",")(1) = "1" Or Split(txtExp(Index).Tag, ",")(3) = "1" Then EnbI
If strT <> vbNullString Then RplTitle Split(txtExp(Index).Tag, ",")(0) = "1" And Split(strT, ",")(0) <> "1" Else: RplTitle vbNullString
If Not bolDebug Then Exit Sub
strT = vbNullString
If Split(txtExp(Index).Tag, ",")(0) = "1" Then strT = "Crucial"
If Split(txtExp(Index).Tag, ",")(1) = "1" Then strT = strT & ", Public"
If Split(txtExp(Index).Tag, ",")(2) = "1" Then strT = strT & ", Array"
If Split(txtExp(Index).Tag, ",")(3) = "1" Then strT = strT & ", Output"
If Left$(strT, 2) = ", " Then strT = Mid$(strT, 3)
txtString(Index).Tag = vbNullString
If bolDebug Then addLog "String {index: " & bytI + 1 & ", number: " & Index + 1 & "} -> " & strT, True
End Sub

Private Sub Form_Load()
'del
'Open App.path & IIf(Right$(App.path, 1) <> "\", "\", vbNullString) & "DotNetCOMRegExLib.dll" For Binary Access Read As #2
'frmPlugins.strRg = "DotNetCOMRegExLib.dll"
'frmPlugins.strC = ",rg1," & vbLf
'bolChk = True
'del
Rnd -Timer * Now
Randomize
strLastPath = vbNullChar
strC = String$(2, """")
ReDim lngProxyPos(0)
Set rh = New cAsyncRequests
If App.LogMode > 0 Then
SetIcon Me.hWnd, "AAA"
Set SystemTray = New clsInTray
SystemTray.hIcon = Me.Icon.handle
strInitD = CurDir$ & IIf(Right$(CurDir$, 1) <> "\", "\", vbNullString)
Else
cmdMintoTray.Enabled = False
SetCurrentDirectoryA App.path
strInitD = App.path & IIf(Right$(App.path, 1) <> "\", "\", vbNullString)
End If
DetC bolMin
If bolMin = vbFalse Then Me.Show Else: cmdMintoTray_Click
'del
'strInitD = "D:\Downloads\"
'ChDrive "D:"
'ChDir "Downloads"
'del
strPl = vbLf
fraH.Tag = vbLf
fraS.Tag = vbLf
cmbField(0).Tag = vbLf & "Cookie" & vbLf & "User-Agent" & vbLf
bytLimit = 19
frmT.bytTOrigin0 = 5
frmPT.bytThreads = 1
frmPT.bytSubThr = 1
frmPT.bytDelay = 1
frmPT.bytTimeout = 20
frmT.bytTOrigin0 = 5
frmT.bytTOrigin1 = 5
frmT.intLogMax = 32767
frmWizard.strUA = DEFUSERAGENT
Dim strT(1) As String
strT(0) = Dir$(App.path & IIf(Right$(App.path, 1) <> "\", "\", vbNullString) & "*.dll", vbHidden)
Do While strT(0) <> vbNullString
If StrComp(strT(0), "dotnetcomregexlib.dll", vbTextCompare) = 0 Then
Dim objT As Object
Set objT = CreateRG
If Not objT Is Nothing Then
On Error GoTo N2
objT.Pattern = "."
objT.Execute "."
Set objT = Nothing
Open App.path & IIf(Right$(App.path, 1) <> "\", "\", vbNullString) & strT(0) For Binary Access Read As #2
frmPlugins.strRg = strT(0)
frmPlugins.strC = ",rg1," & vbLf
bolChk = True
N2:
On Error GoTo 0
End If
Exit Do
End If
strT(0) = Dir$
Loop
If App.LogMode > 0 Then SetCurrentDirectoryA strInitD
If Dir$("Settings.ini", vbHidden) <> vbNullString Then
bolHid = GetAttr("Settings.ini") = vbHidden
On Error GoTo -1
On Error GoTo E
Open "Settings.ini" For Input Access Read As #1
On Error GoTo -1
Dim s() As String, cnt As Integer, l As Boolean, strT1(2) As String, intC(1) As Integer, i As Byte, tmpObj As IPluginInterface
If frmPlugins.strRg <> vbNullString Then cnt = 1
On Error GoTo N
While Not EOF(1)
Line Input #1, strT(0)
strT(0) = Trim$(strT(0))
If InStr(";#[", Left$(strT(0), 1)) = 0 Then
If InStr(strT(0), "=") > 0 Then
strT(1) = Mid$(strT(0), InStr(strT(0), "=") + 1)
If strT(1) <> vbNullString Then
strT(0) = LCase$(Left$(strT(0), InStr(strT(0), "=") - 1))
If IsNumeric(strT(1)) Then
If strT(1) < 0 Then strT(1) = strT(1) * (-1)
Select Case strT(0)
'Case "threads": frmPT.bytThreads = ProcessNumber(strT(1))
Case "subthreads": frmPT.bytSubThr = ProcessNumber(strT(1))
Case "limit": If strT(1) > 7 Or strT(1) < -7 Then bytLimit = ProcessNumber(strT(1) - 1)
Case "timeout": frmPT.bytTimeout = ProcessNumber(strT(1))
Case "sameproxyforeachthread": frmPT.bolSame = strT(1) > 0
Case "skipbadproxies": frmPT.bolSkip = strT(1) > 0
Case "donotstartwithproxy": frmPT.bolNoStartP = strT(1) > 0
Case "donotretry": frmPT.bolNoRetry = strT(1) > 0
Case "donotchangeproxy": frmPT.bolNoChange = strT(1) > 0
Case "delaybetweenretries": frmPT.bytDelay = ProcessNumber(strT(1))
Case "maxretriespercycle": frmPT.bytMaxR = ProcessNumber(strT(1))
Case "maxcycles": frmPT.bytCycles = ProcessNumber(strT(1))
Case "originmax": frmT.bytTOrigin0 = ProcessNumber(strT(1))
Case "debug": bolDebug = strT(1) > 0
Case "exeskiploadingerrors": If strT(1) = "1" Then bolSkipLE = vbUseDefault Else: If strT(1) > 1 Then bolSkipLE = vbTrue
Case "exemoresilent": If strT(1) = "1" Then bolSilent2 = vbUseDefault Else: If strT(1) > 1 Then bolSilent2 = vbTrue
End Select
ElseIf strT(0) = "after" Then
frmT.intAfter = Val(strT(1))
frmT.bolHours = Right$(strT(1), 1) = "h"
ElseIf strT(0) = "originmax" Then
If Right$(strT(1), 1) = "c" Then
frmT.bolColl = True
strT(1) = Left$(strT(1), Len(strT(1)) - 1)
End If
If strT(1) <> vbNullString Then
If Right$(strT(1), 1) = "n" Then
frmT.bolNoEach = True
strT(1) = Left$(strT(1), Len(strT(1)) - 1)
End If
If strT(1) <> vbNullString Then
strT(1) = Trim$(strT(1))
If Left$(strT(1), 1) <> ";" Then frmT.bytTOrigin0 = ProcessNumber(Left$(strT(1), InStr(strT(1) & ";", ";") - 1))
If InStr(strT(1), ";") > 0 Then
frmT.bytTOrigin1 = ProcessNumber(Mid$(strT(1), InStr(strT(1), ";") + 1))
If frmT.bytTOrigin1 = 0 Then frmT.bolColl = True
End If
End If
End If
ElseIf strT(0) = "logoutputmax" Then
strT(1) = Trim$(strT(1))
If Left$(strT(1), 1) <> ";" Then frmT.intLogMax = ProcessNumber(Left$(strT(1), InStr(strT(1) & ";", ";") - 1), True)
If InStr(strT(1), ";") > 0 Then frmT.intOutMax = ProcessNumber(Mid$(strT(1), InStr(strT(1), ";") + 1), True)
ElseIf strT(0) = "output" Then
intC(1) = -1
For i = 0 To 1
intC(0) = intC(1) + 3
intC(1) = FindC(strT(1), intC(0))
If intC(1) = 0 Then Exit For
If i = 0 Then frmT.strTemplate0 = Replace(Replace(Mid$(strT(1), intC(0), intC(1) - intC(0)), "[nl]", vbNewLine), strC, """") Else: frmT.strTemplate1 = Replace(Replace(Mid$(strT(1), intC(0), intC(1) - intC(0)), "[nl]", vbNewLine), strC, """")
Next
ElseIf strT(0) = "pluginsdir" Then
If Dir$(strT(1), vbDirectory) <> vbNullString Then
frmPlugins.strLocation = GetAbsolutePath(strT(1))
frmPlugins.strRL = strT(1)
If Right$(frmPlugins.strLocation, 1) <> "\" Then frmPlugins.strLocation = frmPlugins.strLocation & "\"
End If
ElseIf strT(0) = "loadedplugin" Then
s() = Split(strT(1), "|")
If UBound(s) > 0 Then
strT1(0) = frmPlugins.strLocation & Replace(s(0), frmPlugins.strLocation, vbNullString)
If Dir$(strT1(0), vbHidden) <> vbNullString Then
strT1(0) = GetAbsolutePath(strT1(0))
strT1(1) = Left$(strT1(0), InStrRev(strT1(0), "\"))
strT1(2) = Mid$(strT1(0), Len(strT1(1)) + 1)
frmPlugins.strP = frmPlugins.strP & strT1(1) & "|" & strT1(2)
l = False
On Error GoTo -1
On Error GoTo N1
For i = 1 To UBound(s)
Set tmpObj = CreateObjectEx2(strT1(0), strT1(0), s(i))
If Not tmpObj Is Nothing Then
l = True
Plugins.add tmpObj, strT1(2) & "/" & s(i)
frmPlugins.strC = frmPlugins.strC & "|" & strT1(2) & "/" & s(i) & "|," & TrimComma(tmpObj.Startup(Me)) & "," & vbLf
frmPlugins.strP = frmPlugins.strP & "|" & cnt & "S"
frmPlugins.strPl = frmPlugins.strPl & s(i) & vbLf
'bolChk = True
cnt = cnt + 1
Set tmpObj = Nothing
End If
Next
N1:
If l Then frmPlugins.strP = frmPlugins.strP & "|" & vbLf Else: frmPlugins.strP = Replace(frmPlugins.strP, strT1(1) & "|" & strT1(2), vbNullString)
On Error GoTo -1
On Error GoTo N
End If
End If
ElseIf strT(0) = "defaultwizarduseragent" Then
If strT(1) <> vbNullString Then frmWizard.strUA = strT(1)
Else: If strT(0) = "lastpath" Then If strInitD <> strT(1) Then strLastPath = strT(1) Else: strLastPath = vbNullString
End If
End If
End If
End If
N:
Wend
Close #1
E:
On Error GoTo 0
End If
cmbIndex.AddItem "1"
cmbIndex.ListIndex = 0
cmbGoto(1).AddItem "1", 1
cmbGoto(0).AddItem "1", 1
cmbSign(0).AddItem "<", 1
cmbSign(0).AddItem ">", 1
cmbSign(0).AddItem "=", 0
cmbSign(0).ListIndex = 0
cmbOper(0).ListIndex = 0
cmbThen(1).ListIndex = 0
cmbGoto(1).ListIndex = 1
cmbThen(0).ListIndex = 0
cmbGoto(0).ListIndex = 0
DetC
If CInt(frmPT.bytThreads) + CInt(frmPT.bytSubThr) > 255 Then If frmPT.bytSubThr > frmPT.bytThreads Then frmPT.bytSubThr = frmPT.bytSubThr - frmPT.bytThreads Else: frmPT.bytSubThr = 1
If strLastPath <> vbNullChar Then If strLastPath <> vbNullString Then SetCurrentDirectoryA strLastPath
If Command$ <> vbNullString Then PopulA
If Me.Caption = "UniBot" Then DimP True
If Not cmdMintoTray.Checked Then txtName.SetFocus
If bolDebug Then addLog "Scanning for plugins...", True
If frmPlugins.strLocation = vbNullString Then
frmPlugins.strLocation = strInitD
ScanPlugins frmPlugins.strLocation, True
Else: ScanPlugins frmPlugins.strLocation
End If
lblStatus.Caption = "Idle..."
addLog "Program started."
If Me.Caption <> "UniBot" Then If strCmd <> vbNullString Or bolMin = vbTrue Then cmdStart_Click Else: cmdStart_Click 'del Else
'cmdManager_Click 'del
'frmPlugins.Show vbModal 'del
'frmSB.Show vbModal 'del
'frmEXE.Show vbModal 'del
'frmWizard.Show vbModal 'del
'del
'Debug.Print SaveC(0)
'End
'ChkStr "<aaa%.sad>+%bbb%+'sad'+%ccc%"
'End
'del
End Sub

Private Function DetC(Optional bolMin As VbTriState = vbUseDefault) As Boolean
Dim intC(2) As Integer
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
If bolMin = vbFalse Then
If InStr(" " & Mid$(Command$, intC(0), intC(1) - intC(0)) & " ", " -m ") > 0 Then
bolMin = vbTrue
Exit Function
End If
Else: DetectC intC(0), intC(1)
End If
intC(2) = FindC(Command$, intC(1) + 1) + 1
If intC(2) = 1 Then Exit Do
Loop Until InStr(Mid$(Command$, intC(0), intC(1) - intC(0) - 1), """") > 0
Else: If Command$ <> vbNullString Then If bolMin = vbFalse Then bolMin = InStr(" " & Command$ & " ", " -m ") > 0 Else: DetectC 1, Len(Command$) + 1
End If
End Function

Private Function GetAbsolutePath(strInput As String) As String
Dim fso As FileSystemObject
Set fso = New FileSystemObject
GetAbsolutePath = fso.GetAbsolutePathName(strInput)
Set fso = Nothing
End Function

Private Sub DetectC(intS As Integer, intE As Integer)
Dim strT(1) As String
strT(0) = " " & Mid$(Command$, intS, intE - intS) & " "
Do
If InStr(1, strT(0), " -c ", vbTextCompare) > 0 And Me.Caption = "UniBot" Then
strT(1) = ExtrF("c", intS, strT(0))
If strT(1) <> vbNullString Then
If Dir$(strT(1), vbHidden) <> vbNullString Then
LoadConfig strT(1)
ElseIf InStr(strT(1), ".") = 0 Then strT(1) = strT(1) & ".ini": LoadConfig strT(1)
End If
End If
ElseIf InStr(1, strT(0), " -v ", vbTextCompare) > 0 Then
chkNoSave.Checked = True
strT(0) = Replace(strT(0), " -v", vbNullString)
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
'ElseIf InStr(1, strT(0), " -m ", vbTextCompare) > 0 Then
'cmdMintoTray_Click
'strT(0) = Replace(strT(0), " -m", vbNullString)
ElseIf InStr(1, strT(0), " -w ", vbTextCompare) > 0 Then
frmPT.bolNoStartP = True
strT(0) = Replace(strT(0), " -w", vbNullString)
ElseIf InStr(1, strT(0), " -t ", vbTextCompare) > 0 Then AddVal "t", strT(0)
ElseIf InStr(1, strT(0), " -p ", vbTextCompare) > 0 Then
strT(1) = ExtrF("p", intS, strT(0))
If strT(1) <> vbNullString Then
If InStr(1, strT(1), ":", vbTextCompare) > 0 And InStr(1, strT(1), ".", vbTextCompare) > 0 Then
If RegExpr(ProxyRegex, strT(1)) Then frmPT.strProxy = strT(1)
ElseIf Dir$(strT(1), vbHidden) <> vbNullString Then
frmPT.strProxy = LoadFile(strT(1))
If bolDebug Then addLog "File """ & get_relative_path_to(strT(1)) & """ loaded for proxy list.", True
End If
End If
ElseIf InStr(1, strT(0), " -s ", vbTextCompare) > 0 Then
frmPT.bolSame = True
strT(0) = Replace(strT(0), " -s", vbNullString)
ElseIf InStr(1, strT(0), " -b ", vbTextCompare) > 0 Then
frmPT.bolSkip = True
strT(0) = Replace(strT(0), " -b", vbNullString)
ElseIf InStr(1, strT(0), " -h ", vbTextCompare) > 0 Then AddVal "h", strT(0)
ElseIf InStr(1, strT(0), " -u ", vbTextCompare) > 0 Then AddVal "u", strT(0)
ElseIf InStr(1, strT(0), " -a ", vbTextCompare) > 0 Then AddVal "a", strT(0)
ElseIf InStr(1, strT(0), " -r ", vbTextCompare) > 0 Then AddVal "r", strT(0)
ElseIf InStr(1, strT(0), " -y ", vbTextCompare) > 0 Then AddVal "y", strT(0)
ElseIf InStr(1, strT(0), " -n ", vbTextCompare) > 0 Then
frmPT.bolNoRetry = True
strT(0) = Replace(strT(0), " -n", vbNullString)
ElseIf InStr(1, strT(0), " -g ", vbTextCompare) > 0 Then
frmPT.bolNoChange = True
strT(0) = Replace(strT(0), " -g", vbNullString)
ElseIf InStr(1, strT(0), " -f ", vbTextCompare) > 0 Or InStr(1, strT(0), " -fh ", vbTextCompare) > 0 Then
strT(1) = Split(Split(strT(0), " -")(1), " ")(0)
frmT.bolHours = Right$(strT(1), 1) = "h"
Dim strT1 As String: strT1 = Split(Mid$(strT(0), Len(strT(1)) + 5), " ")(0)
frmT.intAfter = ProcessNumber(strT1, True)
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
Case "t": frmPT.bytTimeout = bytT
Case "h": frmPT.bytThreads = bytT
Case "u": frmPT.bytSubThr = bytT
Case "a": frmPT.bytDelay = bytT
Case "r": frmPT.bytMaxR = bytT
Case "y": frmPT.bytCycles = bytT
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

Private Sub DimP(Optional bolN As Boolean)
If Not bolN Then
ReDim Preserve bytSh(bytLimit)
ReDim Preserve strName(bytLimit)
ReDim Preserve strURLData(bytLimit)
ReDim Preserve bolProxy(1, bytLimit)
ReDim Preserve strWait(1, bytLimit)
ReDim Preserve intGoto(1, bytLimit)
Else
ReDim bytSh(bytLimit)
ReDim strName(bytLimit)
ReDim strURLData(bytLimit)
ReDim bolProxy(1, bytLimit)
ReDim strWait(1, bytLimit)
ReDim intGoto(1, bytLimit)
End If
ReDim strHeaders(bytLimit, (bytLimit + 1) \ 2)
ReDim strStrings(bytLimit, (bytLimit + 1) \ 2)
ReDim strIf(bytLimit, (bytLimit + 1) \ 2)
End Sub

Sub addLog(txt As String, Optional D As Boolean)
txt = "[" & Now & "] " & txt
If D Then txt = "DEBUG: " & txt
If lstLog.ListCount = frmT.intLogMax Then
'strLog = lstLog.list(0) & vbLf & strLog
lstLog.RemoveItem 0
End If
lstLog.AddItem txt
SetListboxScrollbar
lstLog.ListIndex = lstLog.ListCount - 1
lstLog.Text = vbNullString
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode > 0 Or App.LogMode = 0 Then Exit Sub
If Left$(Me.Caption, 1) = "*" And cmdSaveC.Enabled Then
Select Case MsgBox("Save current configuration?", vbYesNoCancel + vbExclamation)
Case vbYes: cmdSaveC_Click
Case vbCancel: GoTo E
End Select
ElseIf Me.Caption <> "UniBot" Then
If MsgBox("Are you sure?", vbExclamation + vbYesNo) = vbNo Then
E:
Cancel = 1
Exit Sub
End If
End If
End Sub

Private Sub lstLog_DblClick()
On Error Resume Next
Clipboard.Clear
Clipboard.SetText lstLog.Text
End Sub

Private Sub txtA_Validate(Index As Integer, Cancel As Boolean)
If Left$(lblStatus.Caption, 5) = "Shift" Or Left$(lblStatus.Caption, 4) = "Load" Or Left$(lblStatus.Caption, 5) = "Remov" Then Exit Sub
txtA(Index).Text = Trim$(Replace(Replace(txtA(Index).Text, vbLf, vbNullString), vbCr, vbNullString))
If Index = 0 Then
If txtA(0).Text = "[clear]" Then
If txtA.count = 1 Then Exit Sub
Dim i As Byte, bolT As Boolean
txtA(0).Text = vbNullString
txtB(0).Text = vbNullString
For i = 1 To txtA.count - 1
txtA(i).Text = vbNullString
txtB(i).Text = vbNullString
If strIf(bytI, i) <> vbNullString And Not bolT Then bolT = True
strIf(bytI, i) = vbNullString
Next
cmdAdd(2).Enabled = False
'If bolDebug And bolT Then addLog "If {index: " & bytI + 1 & ", all} ->", True
'RplTitle
End If
End If
If AddIf(Index, True) Then Cancel = True
End Sub

Private Sub txtB_Validate(Index As Integer, Cancel As Boolean)
If Left$(lblStatus.Caption, 5) = "Shift" Or Left$(lblStatus.Caption, 4) = "Load" Or Left$(lblStatus.Caption, 5) = "Remov" Then Exit Sub
txtB(Index).Text = Trim$(Replace(Replace(txtB(Index).Text, vbLf, vbNullString), vbCr, vbNullString))
If AddIf(Index, True, True) Then Cancel = True
End Sub

Private Function AddIf(Index As Integer, Optional bolT As Boolean, Optional bolB As Boolean) As Boolean
Dim strI As String, bolR As Boolean
If bolT Then
Dim i As Byte, j As Byte, bytS As Byte
If Index < txtA.count - 1 Then
For i = Index + 1 To txtA.count - 1
If txtA(i).Text <> vbNullString Or txtB(i).Text <> vbNullString Then
bytS = i
Exit For
End If
Next
ElseIf txtA(Index).Text <> vbNullString Or txtB(Index).Text <> vbNullString Then bytS = Index
End If
If Index > 0 Or bytS > 0 Then
If Index > 0 And bytS > 0 Then
If txtA(Index).Text = vbNullString And txtB(Index).Text = vbNullString And cmbOper(Index - 1).ListIndex = 1 Then
HScroll1.Value = HScroll1.LargeChange * Index
If MsgBox("Shift items?", vbYesNo + vbQuestion) = vbNo Then GoTo Skip
End If
End If
If txtA(Index).Text = vbNullString And txtB(Index).Text = vbNullString Then ' Or txtB(0).Text = vbNullString And txtA(0).Text = vbNullString
cmdAdd(2).Enabled = False
'If txtA(Index).Text = vbNullString And txtB(Index).Text = vbNullString Then
i = Index
If bytSh(bytI) = i And bytS = 0 Then
Do While bytSh(bytI) > 0
If txtA(bytSh(bytI)).Text <> vbNullString Or txtB(bytSh(bytI)).Text <> vbNullString Then Exit Do
bytSh(bytI) = bytSh(bytI) - 1
Loop
End If
'Else
'i = 0
'Index = 0
'HScroll1.value = 0
'End If
j = bytS
If j > 0 Then
If bytSh(bytI) > 0 Then bytSh(bytI) = bytSh(bytI) - (j - i)
Do While j <= txtA.count - 1
ReplIf i, j
i = i + 1
j = j + 1
Loop
End If
Else
If Index > 0 Then
If Index > bytSh(bytI) Then bytSh(bytI) = Index 'If txtA(Index).Text <> vbNullString Or txtB(Index).Text <> vbNullString Then
i = Index
j = Index
Dim bolO As Boolean
Do
i = i - 1
If i > 0 Then
If cmbOper(i - 1).ListIndex = 0 Then bolO = True Else: bolO = False
Else: bolO = True
End If
If bolO And txtA(i).Text = vbNullString And txtB(i).Text = vbNullString Then
If i > bytSh(bytI) Then bytSh(bytI) = i
ReplIf i, j
HScroll1.Value = HScroll1.LargeChange * i
If Index = txtA.count - 1 Then cmdAdd(2).Enabled = False
Index = i
Exit Do
End If
Loop Until i = 0
End If
End If
End If
If Index = 0 Then
If bytI <> cmbIndex.ListCount - 1 Then 'And cmbIndex.Enabled
If txtB(0).Text = vbNullString And txtA(0).Text = vbNullString Then
If Not Filled(bytI, 2) Then
If RemI Then AddIf = True ': cmbIndex.Enabled = False: cmdStart.Enabled = False
'If bolDebug Then
'If txtA(0).Text <> vbNullString Or txtB(0).Text <> vbNullString Then
'txtA(0).Tag = "."
'bolR = True
'ElseIf txtA(0).Tag = "." Then
'txtA(0).Tag = vbNullString
'bolR = True
'End If
'If bolR Then GoTo P Else: Exit Function
'End If
End If
End If
End If
End If
End If
Skip:
If txtA(Index).Text <> vbNullString Or txtB(Index).Text <> vbNullString Then
strI = txtA(Index).Text & vbLf & cmbSign(Index).ListIndex & vbLf & txtB(Index).Text
If Index > 0 Then strI = strI & vbLf & cmbOper(Index - 1).ListIndex
If strI <> strIf(bytI, Index) Then
strIf(bytI, Index) = strI
If cmbIndex.ListCount - 1 = bytI Then AddI 'Or Not cmbIndex.Enabled
If strIf(bytI, txtA.count - 1) <> vbNullString Then cmdAdd(2).Enabled = True
bolR = True
End If
ElseIf strIf(bytI, Index) <> vbNullString Then
strIf(bytI, Index) = vbNullString
If strIf(bytI, txtA.count - 1) <> vbNullString Then cmdAdd(2).Enabled = True
bolR = True
End If
If bolR Then
EnbI
If bolB Then
If txtB(Index).Text <> vbNullString Then RplTitle ChkStr(txtB(Index).Text, 1) Else: RplTitle vbNullString
Else: If txtA(Index).Text <> vbNullString Then RplTitle ChkStr(txtA(Index).Text, 1) Else: RplTitle vbNullString
End If
If Not bolDebug Then Exit Function
Dim strO As String
If Index > 0 Then strO = cmbOper(Index - 1).Text & " "
strO = strO & txtA(Index).Text & " " & cmbSign(Index).Text & " " & txtB(Index).Text
addLog "If {index: " & bytI + 1 & ", number: " & Index + 1 & "} -> " & strO, True
End If
End Function

Private Sub ReplIf(ByVal i As Byte, j As Byte)
txtA(i).Text = txtA(j).Text
bolAl = True
cmbSign(i).ListIndex = cmbSign(j).ListIndex
txtB(i).Text = txtB(j).Text
Dim strI As String
strI = txtA(j).Text & vbLf & cmbSign(j).Text & vbLf & txtB(j).Text
If j > 0 Then strI = strI & vbLf & cmbOper(j - 1).Text
strIf(bytI, i) = strI
txtA(j).Text = vbNullString
cmbSign(j).ListIndex = 0
bolAl = False
txtB(j).Text = vbNullString
strIf(bytI, j) = vbNullString
End Sub

Function RemI(Optional bolI As Boolean, Optional bolA As Boolean) As Boolean 'Optional bytC As Byte
Dim bytI1 As Byte
'If bytC = 0 Then
If Not bolA Then
If MsgBox("This action will completely remove this index." & vbNewLine & "Are you sure that you want to continue?", vbExclamation + vbYesNo) = vbNo Then
If Not bolI Then
EnbC True
cmdI.Enabled = False
End If
RemI = True
Exit Function
End If
End If
If Me.Caption <> "UniBot" Then If Not bolA And Left$(Me.Caption, 1) <> "*" Then Me.Caption = "*" & Me.Caption
bytI1 = bytI
lblStatus.Caption = "Shifting indexes from " & bytI1 + 1 & "..."
lblStatus.Refresh
'disF
Screen.MousePointer = 11
'Else: bytI1 = bytC - 1
'End If
If InStr(cmdOpt(0).Tag, "-" & bytI1 & "-") > 0 Then cmdOpt(0).Tag = Replace(cmdOpt(0).Tag, "-" & bytI1 & "-" & Split(Split(cmdOpt(0).Tag, "-" & bytI1 & "-")(1), vbLf)(0) & vbLf, vbNullString)
If cmbIndex.ListIndex < cmbIndex.ListCount - 1 Then
Dim i As Byte, j As Byte
i = bytI1
j = bytI1 + 1
Do While j <= cmbIndex.ListCount - 1
ChngI i, j
i = i + 1
j = j + 1
Loop
'If bytC > 0 Then Exit Sub
cmbIndex.RemoveItem cmbIndex.ListCount - 1
cmbGoto(0).RemoveItem cmbGoto(0).ListCount - 1
cmbGoto(1).RemoveItem cmbGoto(1).ListCount - 1
Else
ChngI cmbIndex.ListCount - 1, cmbIndex.ListCount - 1
bolL = False
End If
cmbIndex.Enabled = True
EnbC
If cmbIndex.ListCount = 1 Then
Me.Caption = "UniBot"
bolChk = False
strPl = vbLf
cmdNew.Enabled = False
EnbC True
cmdR.Enabled = False
cmdStart.Tag = vbNullString
cmdStart.Enabled = False
Else
If cmbIndex.ListCount = bytLimit And Filled(cmbIndex.ListCount - 1) Then AddI
cmdStart.Enabled = True
End If
cmbIndex_Click
If bolDebug Then addLog "Index " & bytI1 + 1 & " removed.", True
EnbIn
If frmManager.bolL Then frmManager.Tag = " "
End Function

Sub ChngI(i As Byte, j As Byte, Optional bolA As Boolean)
strName(i) = strName(j)
strName(j) = vbNullString
strURLData(i) = strURLData(j)
strURLData(j) = vbNullString
'Dim s() As String
Dim a As Byte, b As Byte
For a = 0 To UBound(strHeaders, 2)
If strHeaders(j, a) = vbNullString Then
If strHeaders(i, a) <> vbNullString Then
For b = a To UBound(strHeaders, 2)
If strHeaders(i, b) = vbNullString Then Exit For
's() = Split(strHeaders(i, b), vbLf)
'frmH.Tag = Replace(frmH.Tag, vbLf & " " & s(0) & " " & vbLf & s(1) & vbLf, vbLf)
strHeaders(i, b) = vbNullString
Next
End If
Exit For
End If
's() = Split(strHeaders(i, a), vbLf)
'frmH.Tag = Replace(frmH.Tag, vbLf & " " & s(0) & " " & vbLf & s(1) & vbLf, vbLf)
strHeaders(i, a) = strHeaders(j, a)
strHeaders(j, a) = vbNullString
Next
For a = 0 To UBound(strStrings, 2)
If strStrings(j, a) = vbNullString Then
If strStrings(i, a) <> vbNullString Then
For b = a To UBound(strStrings, 2)
If strStrings(i, b) = vbNullString Then Exit For
RemPu i, b
strStrings(i, b) = vbNullString
Next
End If
Exit For
End If
RemPu i, b
strStrings(i, a) = strStrings(j, a)
strStrings(j, a) = vbNullString
Next
For a = 0 To UBound(strIf, 2)
If strIf(j, a) = vbNullString Then
If strIf(i, a) <> vbNullString Then
For b = a To UBound(strIf, 2)
If b > bytSh(i) And strIf(i, b) = vbNullString Then Exit For
strIf(i, b) = vbNullString
Next
End If
Exit For
End If
strIf(i, a) = strIf(j, a)
strIf(j, a) = vbNullString
Next
bytSh(i) = 0
For a = 0 To 1
strWait(a, i) = strWait(a, j)
strWait(a, j) = vbNullString
bolProxy(a, i) = bolProxy(a, j)
bolProxy(a, j) = False
If Not bolA Then
If intGoto(a, j) - 2 >= 0 Then
If intGoto(a, j) - 2 > bytI Then
intGoto(a, i) = intGoto(a, j) - 1
ElseIf intGoto(a, j) - 2 < bytI Then intGoto(a, i) = intGoto(a, j)
Else: intGoto(a, i) = 0
End If
End If
Else: If intGoto(a, j) - 2 >= bytI Then intGoto(a, i) = intGoto(a, j) + 1 Else: intGoto(a, i) = intGoto(a, j)
End If
intGoto(a, j) = 0
Next
If InStr(cmdOpt(0).Tag, "-" & j & "-") > 0 Then cmdOpt(0).Tag = Replace(cmdOpt(0).Tag, "-" & j & "-" & Split(Split(cmdOpt(0).Tag, "-" & j & "-")(1), vbLf)(0) & vbLf, vbNullString, , 1)
cmdOpt(0).Tag = Replace(cmdOpt(0).Tag, "-" & j & "-", "-" & i & "-", , 1)
End Sub

Private Sub RemPu(i As Byte, b As Byte)
If strStrings(i, b) = vbNullString Then Exit Sub
If Split(strStrings(i, b), vbLf)(2) = vbNullString Then Exit Sub
Dim s() As String
s() = Split(Split(strStrings(i, b), vbLf)(2), ",")
If s(1) = "1" Then CheckPublic Split(strStrings(i, b), vbLf)(0)
End Sub

Sub SLInd(i As Byte, Optional bolS As Boolean)
Static strN As String, strU As String, strH() As String, strS() As String, strI() As String, bytS As Byte, bolP(1) As Boolean, strW(1) As String, intG(1) As Integer, bytO As Byte
Dim a As Byte
If Not bolS Then
strN = strName(i)
strU = strURLData(i)
Do While strHeaders(i, a) <> vbNullString
ReDim Preserve strH(a)
strH(a) = strHeaders(i, a)
a = a + 1
If a > UBound(strHeaders, 2) Then Exit Do
Loop
a = 0
Do While strStrings(i, a) <> vbNullString
ReDim Preserve strS(a)
strS(a) = strStrings(i, a)
a = a + 1
If a > UBound(strStrings, 2) Then Exit Do
Loop
a = 0
Do While strIf(i, a) <> vbNullString
ReDim Preserve strI(a)
strI(a) = strIf(i, a)
a = a + 1
If a > UBound(strIf, 2) Then Exit Do
Loop
bytS = bytSh(i)
For a = 0 To 1
strW(a) = strWait(a, i)
bolP(a) = bolProxy(a, i)
intG(a) = intGoto(a, i)
Next
If InStr(cmdOpt(0).Tag, "-" & i & "-") > 0 Then bytO = Split(Split(cmdOpt(0).Tag, "-" & i & "-")(1), vbLf)(0)
Else
strName(i) = strN
strN = vbNullString
strURLData(i) = strU
strU = vbNullString
On Error GoTo N
For a = 0 To UBound(strH)
strHeaders(i, a) = strH(a)
Next
Erase strH
N:
If strHeaders(i, a) <> vbNullString Then
Do While strHeaders(i, a) <> vbNullString
strHeaders(i, a) = vbNullString
a = a + 1
If a > UBound(strHeaders, 2) Then Exit Do
Loop
End If
On Error GoTo -1
On Error GoTo N1
For a = 0 To UBound(strS)
strStrings(i, a) = strS(a)
If Split(Split(strStrings(i, a), vbLf)(2) & ",", ",")(1) = "1" Then CheckPublic Split(strStrings(i, a), vbLf)(0), True, True
Next
Erase strS
N1:
If strStrings(i, a) <> vbNullString Then
Do While strStrings(i, a) <> vbNullString
strStrings(i, a) = vbNullString
a = a + 1
If a > UBound(strStrings, 2) Then Exit Do
Loop
End If
On Error GoTo -1
On Error GoTo N2
For a = 0 To UBound(strI)
strIf(i, a) = strI(a)
Next
Erase strI
N2:
If strIf(i, a) <> vbNullString Then
For a = a To bytSh(i)
strIf(i, a) = vbNullString
Next
End If
bytSh(i) = bytS
For a = 0 To 1
strWait(a, i) = strW(a)
strW(a) = vbNullString
bolProxy(a, i) = bolP(a)
intGoto(a, i) = intG(a)
Next
If bytO > 0 Then cmdOpt(0).Tag = cmdOpt(0).Tag & "-" & i & "-" & bytO & vbLf
End If
End Sub

Private Sub txtExp_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> vbKeyReturn Then Exit Sub
If Left$(txtExp(Index).Text, 7) = "[build]" Then
KeyAscii = 0
frmSB.Caption = "[String builder] " & txtString(Index).Text
frmSB.Tag = Index
frmSB.txtInput.Text = Mid$(txtExp(Index).Text, 8)
frmSB.Show vbModal
txtExp(Index).SetFocus
ElseIf txtExp(Index).Text = "[file]" Then
KeyAscii = 0
txtExp(Index).Text = CommDlg(, txtString(Index).Text, , 1)
If txtExp(Index).Text = vbNullString Then Exit Sub
txtExp(Index).Text = get_relative_path_to(txtExp(Index).Text)
txtExp(Index).Text = "<" & txtExp(Index).Text & ">"
txtExp(Index).SelStart = Len(txtExp(Index).Text)
Else: txtExp_Validate Index, False
End If
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
If Left$(lblStatus.Caption, 5) = "Shift" Or Left$(lblStatus.Caption, 4) = "Load" Or Left$(lblStatus.Caption, 5) = "Remov" Then Exit Sub
txtName.Text = Trim$(Replace(Replace(txtName.Text, vbLf, vbNullString), vbCr, vbNullString))
If strName(bytI) = txtName.Text Then Exit Sub
strName(bytI) = txtName.Text
RplTitle
If bolDebug Then addLog "Name {index: " & bytI + 1 & "}: " & txtName.Text, True
End Sub

Private Sub txtExp_Validate(Index As Integer, Cancel As Boolean)
If Left$(lblStatus.Caption, 5) = "Shift" Or Left$(lblStatus.Caption, 4) = "Load" Or Left$(lblStatus.Caption, 5) = "Remov" Then Exit Sub
txtExp(Index).Text = Trim$(Replace(Replace(txtExp(Index).Text, vbLf, vbNullString), vbCr, vbNullString))
'If txtString(Index).Text = vbNullString And txtExp(Index).Text = vbNullString Then Exit Sub
If AddToH(FindH(Index, True), True, True) Then Cancel = True
End Sub

Private Sub txtOutput_KeyPress(KeyAscii As Integer)
If KeyAscii <> 1 Then Exit Sub
txtOutput.SelStart = 0
txtOutput.SelLength = Len(txtOutput.Text)
End Sub

Private Sub txtString_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> vbKeyReturn Then Exit Sub
If Index > 0 Then Exit Sub
If txtString(0).Text <> "[clear]" Then Exit Sub
If strStrings(bytI, 1) = vbNullString Then Exit Sub
If Filled(bytI, 3) Then
Dim i As Byte
For i = 0 To UBound(strStrings, 2)
If strStrings(bytI, i) = vbNullString Then Exit For
strStrings(bytI, i) = vbNullString
txtString(i).Text = vbNullString
txtExp(i).Text = vbNullString
txtExp(i).Tag = vbNullString
cmdOpt(i).Enabled = False
Next
cmdAdd(1).Enabled = False
If InStr(cmdOpt(0).Tag, "-" & bytI & "-") > 0 Then cmdOpt(0).Tag = Replace(cmdOpt(0).Tag, "-" & bytI & "-" & Split(Split(cmdOpt(0).Tag, "-" & bytI & "-")(1), vbLf)(0) & vbLf, vbNullString)
Else: RemI
End If
End Sub

Private Sub txtString_Validate(Index As Integer, Cancel As Boolean)
If Left$(lblStatus.Caption, 5) = "Shift" Or Left$(lblStatus.Caption, 4) = "Load" Or Left$(lblStatus.Caption, 5) = "Remov" Then Exit Sub
txtString(Index).Text = Trim$(Replace(Replace(Replace(Replace(Replace(txtString(Index).Text, "%", ""), "{", ""), "}", ""), vbLf, vbNullString), vbCr, vbNullString))
Index = FindH(Index, True)
If txtString(Index).Text <> vbNullString Then If ChkDup(txtString, Index) Then Cancel = True: Exit Sub
Cancel = AddToH(Index, True)
End Sub

Private Function ChkDup(objF As Object, Index As Integer) As Boolean
Dim i As Byte
For i = 0 To objF.count - 1
If i <> Index Then
If objF(i).Text <> vbNullString Then
If objF(i).Text = objF(Index).Text Then
MsgBox "Field with same name already exists!", vbExclamation
ChkDup = True
objF(Index).SelStart = 0
objF(Index).SelLength = Len(objF(Index).Text)
objF(Index).SetFocus
Exit Function
End If
End If
End If
Next
End Function

Private Sub txtURL_Validate(Cancel As Boolean)
If Left$(lblStatus.Caption, 5) = "Shift" Or Left$(lblStatus.Caption, 4) = "Load" Or Left$(lblStatus.Caption, 5) = "Remov" Then Exit Sub
txtURL.Text = Trim$(Replace(Replace(txtURL.Text, vbLf, vbNullString), vbCr, vbNullString))
If txtURL.Text = Split(strURLData(bytI) & vbLf, vbLf)(0) Then
If txtURL.Text = vbNullString And txtURL.Tag <> vbNullString Then
txtURL.Tag = vbNullString
If bolDebug Then addLog "URL -> ()", True
End If
Exit Sub
End If
Dim bolT As Boolean
If ChkURL(txtURL.Text, bolT) Then
If bytI <> cmbIndex.ListCount - 1 Then 'And cmbIndex.Enabled
If Not Filled(bytI, 1) Then
If RemI Then
txtURL.SelStart = 0
txtURL.SelLength = Len(txtURL.Text)
Cancel = True
End If
Exit Sub
Else: RplTitle vbNullString
End If
End If
strURLData(bytI) = vbNullString 'vbLf & Split(strURLData(bytI), vbLf)(1)
If bolDebug And txtURL.Tag <> vbNullString Then
txtURL.Tag = vbNullString
addLog "URL -> ()", True
End If
Else
If cmbIndex.ListCount - 1 = bytI Then AddI 'Or Not cmbIndex.Enabled
If txtData.Text <> vbNullString Then strURLData(bytI) = txtURL.Text & vbLf & Split(strURLData(bytI) & vbLf, vbLf)(1) Else: strURLData(bytI) = txtURL.Text
EnbI
RplTitle bolT
If bolDebug Then addLog "URL {index: " & bytI + 1 & "} -> " & txtURL.Text, True
End If
End Sub

Private Sub txtData_Validate(Cancel As Boolean)
If Left$(lblStatus.Caption, 5) = "Shift" Or Left$(lblStatus.Caption, 4) = "Load" Or Left$(lblStatus.Caption, 5) = "Remov" Then Exit Sub
txtData.Text = Trim$(Replace(Replace(txtData.Text, vbLf, vbNullString), vbCr, vbNullString))
If txtData.Text = Split(strURLData(bytI) & vbLf, vbLf)(1) Then Exit Sub
If txtData.Text <> vbNullString Then strURLData(bytI) = Split(strURLData(bytI) & vbLf, vbLf)(0) & vbLf & txtData.Text Else: strURLData(bytI) = Split(strURLData(bytI) & vbLf, vbLf)(0)
If txtData.Text <> vbNullString Then
If Left$(txtData.Text, 1) = "[" And Right$(txtData.Text, 1) = "]" And InStr(txtData.Text, ":") > 0 Then
If ChkDat(txtData.Text) Then RplTitle vbNullString
Else: RplTitle ChkStr(txtData.Text, 1)
End If
Else: RplTitle vbNullString
End If
If bolDebug Then addLog "POST {index: " & bytI + 1 & "} -> " & txtData.Text, True
End Sub

Private Sub RplTitle(Optional varT As Variant)
If Me.Caption = "UniBot" Then Exit Sub
If Left$(Me.Caption, 1) <> "*" Then Me.Caption = "*" & Me.Caption
If Not cmbIndex.Enabled Then Exit Sub
EnbC
If VarType(varT) = vbBoolean Then
If strPl <> vbLf Then
bolChk = True
strPl = vbLf
ElseIf Not varT And Not bolChk And bytI <> cmbIndex.ListCount - 1 + CInt(bytI = bytLimit) Then bolChk = True
End If
Else: If Not IsMissing(varT) And strPl <> vbLf And cmbIndex.ListCount = 2 Then ChkInd
End If
End Sub

Private Sub EnbC(Optional bolE As Boolean)
If Not bolE Then
If Left$(Me.Caption, 1) = "*" Then cmdSaveC.Enabled = True Else: cmdSaveC.Enabled = False
cmdShortcut.Enabled = True
cmdManager.Enabled = True
cmdMake.Enabled = True
cmdI.Enabled = True
cmdR.Enabled = True
Else
cmdSaveC.Enabled = False
cmdShortcut.Enabled = False
cmdManager.Enabled = False
cmdMake.Enabled = False
End If
End Sub

Private Sub EnbI()
cmbIndex.Enabled = True
cmdStart.Enabled = True
End Sub

Private Sub ChkInd(Optional intI As Integer = -1)
Dim strT As String, a As Byte
If intI = -1 Then a = 0 Else: a = intI
If strURLData(a) <> vbNullString Then
strT = Split(strURLData(a), vbLf)(0)
If strT <> vbNullString Then
If intI = -1 Then
If Not ChkStr(strT, 1) Then Exit Sub
Else: If bolDebug Then ChkStr strT, 2, a Else: ChkStr strT, 2
End If
End If
If InStr(strURLData(a), vbLf) > 0 Then
strT = Split(strURLData(a), vbLf)(1)
If Left$(strT, 1) = "[" And Right$(strT, 1) = "]" And InStr(strT, ":") > 0 Then
If Not ChkDat(strT, Not CBool(intI), a) And intI = -1 Then Exit Sub
ElseIf intI = -1 Then
If Not ChkStr(strT, 1) Then Exit Sub
Else: If bolDebug Then ChkStr strT, 2, a, 1 Else: ChkStr strT, 2
End If
End If
End If
Dim i As Byte
Do While strHeaders(a, i) <> vbNullString
strT = Split(strHeaders(a, i), vbLf)(1)
If intI = -1 Then
If Not ChkStr(strT, 1) Then Exit Sub
Else: If bolDebug Then ChkStr strT, 2, a, 2 Else: ChkStr strT, 2
End If
i = i + 1
If UBound(strHeaders, 2) < i Then Exit Do
Loop
i = 0
Do While strStrings(a, i) <> vbNullString
strT = Split(strStrings(a, i), vbLf)(1)
If intI = -1 Then
If Not ChkStr(strT, 1) Then Exit Sub
Else: If bolDebug Then ChkStr strT, 2, a, 3, i, Split(Split(strStrings(a, i), vbLf)(2) & ",", ",")(0) <> "1" Else: ChkStr strT, 2
End If
i = i + 1
If UBound(strStrings, 2) < i Then Exit Do
Loop
For i = 0 To UBound(strIf, 2)
If strIf(a, i) <> vbNullString Then
strT = Split(strIf(a, i), vbLf)(0)
If intI = -1 Then
If Not ChkStr(strT, 1) Then Exit Sub
Else: If bolDebug Then ChkStr strT, 2, a, 4, i Else: ChkStr strT, 2
End If
strT = Split(strIf(a, i), vbLf)(2)
If intI = -1 Then
If Not ChkStr(strT, 1) Then Exit Sub
Else: If bolDebug Then ChkStr strT, 2, a, 4, i Else: ChkStr strT, 2
End If
ElseIf i > bytSh(i) Then Exit For
End If
Next
If strWait(0, a) <> vbNullString Then
If intI = -1 Then
If Not ChkStr(strWait(0, a), 1) Then Exit Sub
Else: If bolDebug Then ChkStr strWait(0, a), 2, a, 5 Else: ChkStr strWait(0, a), 2
End If
End If
If strWait(1, a) <> vbNullString Then
If intI = -1 Then
If Not ChkStr(strWait(1, a), 1) Then Exit Sub
Else: If bolDebug Then ChkStr strWait(1, a), 2, a, 5, 1 Else: ChkStr strWait(1, a), 2
End If
End If
bolChk = False
If intI = -1 Then strPl = vbLf
End Sub

Private Function ChkDat(strT As String, Optional bolP As Boolean, Optional bytT As Byte) As Boolean
Dim intC(1) As Long, bolT As Boolean, strT1 As String, strC As String
intC(0) = 2
Do While intC(0) < Len(strT)
bolT = Not bolT
If bolT Then strC = ":" Else: strC = ";"
intC(1) = FindSep(strT, intC(0), strC)
If intC(1) = 0 Then intC(1) = Len(strT)
strT1 = Mid$(strT, intC(0), intC(1) - intC(0))
If Not bolP Then
If Not ChkStr(strT1, 1) Then Exit Function
Else: If bolDebug Then ChkStr strT1, 2, bytT, 1 Else: ChkStr strT1, 2
End If
intC(0) = intC(1) + 1
Loop
ChkDat = True
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

Private Sub cmbGoto_Click(Index As Integer)
If lblStatus.Caption = "Starting..." Or Left$(lblStatus.Caption, 5) = "Shift" Or Left$(lblStatus.Caption, 4) = "Load" Or Left$(lblStatus.Caption, 5) = "Remov" Then Exit Sub
If Index = 0 Then
If cmbGoto(Index).ListIndex = 0 Then
If intGoto(0, bytI) = 0 Then Exit Sub
Else: If intGoto(0, bytI) = cmbGoto(Index).ListIndex + 1 Then Exit Sub
End If
If cmbGoto(Index).ListIndex <> 0 Then intGoto(Index, bytI) = cmbGoto(Index).ListIndex + 1 Else: intGoto(Index, bytI) = 0
Else: If intGoto(1, bytI) <> cmbGoto(Index).ListIndex + 1 Then intGoto(Index, bytI) = cmbGoto(Index).ListIndex + 1 Else: Exit Sub
End If
RplTitle
If bolDebug Then addLog "Go to {index: " & bytI + 1 & ", number " & Index + 1 & "}: " & intGoto(Index, bytI), True
End Sub

Sub AddI()
If bolL Then Exit Sub
If cmbIndex.ListCount - 1 = bytLimit Then If Filled(cmbIndex.ListCount - 1) Then If Not AddL Then Exit Sub
If Not cmdNew.Enabled Then
Me.Caption = "* - UniBot"
cmdNew.Enabled = True
End If
Dim bytIndex As Integer: bytIndex = cmbIndex.ListCount + 1
cmbIndex.AddItem bytIndex
cmbGoto(0).AddItem bytIndex
cmbGoto(1).AddItem bytIndex
If bytIndex = 2 Then cmdStart.Enabled = True
End Sub

Private Sub cmbField_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If cmbField(Index).Text = vbNullString Or cmbField(Index).SelLength <> Len(cmbField(Index).Text) Then
cmbField_Validate Index, False
Exit Sub
End If
cmbField(0).Tag = Replace(cmbField(0).Tag, vbLf & cmbField(Index).Text & vbLf, vbLf)
cmbField(Index).Text = vbNullString
End Sub

Private Sub cmbField_Validate(Index As Integer, Cancel As Boolean)
If lblStatus.Caption = "Starting..." Or Left$(lblStatus.Caption, 5) = "Shift" Or Left$(lblStatus.Caption, 4) = "Load" Or Left$(lblStatus.Caption, 5) = "Remov" Then Exit Sub
cmbField(Index).Text = Trim$(Replace(Replace(cmbField(Index).Text, vbLf, vbNullString), vbCr, vbNullString))
Index = FindH(Index)
If cmbField(Index).Text <> vbNullString Then
If cmbField(Index).ListIndex = -1 Then If ChkDup(cmbField, Index) Then Cancel = True: Exit Sub
If txtValue(Index).Text = vbNullString And InStr(fraH.Tag, vbLf & " " & cmbField(Index).Text & " " & vbLf) > 0 Then
With txtValue(Index)
.Text = Split(Split(fraH.Tag, vbLf & " " & cmbField(Index).Text & " " & vbLf)(1), vbLf)(0)
.SetFocus
.SelStart = 0
.SelLength = Len(.Text)
End With
End If
End If
AddF Index
If cmbField(Index).Text <> vbNullString Then If InStr(cmbField(0).Tag, vbLf & cmbField(Index).Text & vbLf) = 0 Then cmbField(0).Tag = cmbField(0).Tag & cmbField(Index).Text & vbLf
AddToH Index
End Sub

Private Function FindH(Index As Integer, Optional bolE As Boolean) As Byte
Dim a As Byte, b As Byte, i As Byte
If Not bolE Then
If cmbField(Index).Text <> vbNullString And txtValue(Index).Text <> vbNullString Then
For i = 0 To Index
If cmbField(i).Text = vbNullString Or txtValue(i).Text = vbNullString Then
ReplH i, CByte(Index)
FindH = i
Exit Function
End If
Next
ElseIf cmbField(Index).Text = vbNullString And txtValue(Index).Text = vbNullString Then
a = Index
b = Index + 1
Do While b <= cmbField.count - 1
ReplH a, b
a = b
b = b + 1
Loop
End If
ElseIf txtString(Index).Text <> vbNullString And txtExp(Index).Text <> vbNullString Then
For i = 0 To Index
If txtString(i).Text = vbNullString Or txtExp(i).Text = vbNullString Then
ReplH i, CByte(Index), True
FindH = i
Exit Function
End If
Next
ElseIf txtString(Index).Text = vbNullString And txtExp(Index).Text = vbNullString Then
a = Index
b = Index + 1
Do While b <= txtExp.count - 1
ReplH a, b, True
a = b
b = b + 1
Loop
End If
FindH = Index
End Function

Private Sub ReplH(ByVal a As Byte, ByVal b As Byte, Optional bolE As Boolean)
If Not bolE Then
cmbField(a).Text = cmbField(b).Text
txtValue(a).Text = txtValue(b).Text
strHeaders(bytI, a) = strHeaders(bytI, b)
cmbField(b).Text = vbNullString
txtValue(b).Text = vbNullString
strHeaders(bytI, b) = vbNullString
Else
txtString(a).Text = txtString(b).Text
txtExp(a).Text = txtExp(b).Text
txtExp(a).Tag = txtExp(b).Tag
If txtString(a).Text <> vbNullString And txtExp(a).Text <> vbNullString Then cmdOpt(a).Enabled = True
strStrings(bytI, a) = strStrings(bytI, b)
txtString(b).Text = vbNullString
txtExp(b).Text = vbNullString
txtExp(b).Tag = vbNullString
cmdOpt(b).Enabled = False
strStrings(bytI, b) = vbNullString
End If
End Sub

Private Sub AddF(Optional Index As Integer, Optional bytS As Byte)
Dim s() As String, strI As String, strT As String
strI = cmbField(0).Tag
s() = Split(strI, vbLf)
Dim i As Byte
If bytS <> 1 Then
For i = 0 To cmbField.count - 1
If cmbField(i).Text <> vbNullString Then
strI = Replace(strI, vbLf & cmbField(i).Text & vbLf, vbLf)
s() = Split(strI, vbLf)
End If
Next
End If
If bytS <> 2 Then
Dim j As Byte, b As Byte
For i = 0 To cmbField.count - 1
If bytS = 0 And i = Index Then GoTo N
strT = cmbField(i).Text
cmbField(i).Clear
If strI <> vbLf Then
For j = 1 To UBound(s()) - 1
If bytS = 1 Then
For b = 0 To UBound(strHeaders, 2)
If strHeaders(bytI, b) = vbNullString Then Exit For
If Split(strHeaders(bytI, b), vbLf)(0) = s(j) Then GoTo N1
Next
End If
cmbField(i).AddItem s(j)
N1:
Next
End If
cmbField(i).Text = strT
N:
Next
ElseIf strI <> vbLf Then
For i = 1 To UBound(s()) - 1
cmbField(Index).AddItem s(i)
Next
End If
End Sub

Private Sub txtValue_Validate(Index As Integer, Cancel As Boolean)
If Left$(lblStatus.Caption, 5) = "Shift" Or Left$(lblStatus.Caption, 4) = "Load" Or Left$(lblStatus.Caption, 5) = "Remov" Then Exit Sub
txtValue(Index).Text = Trim$(Replace(Replace(txtValue(Index).Text, vbLf, vbNullString), vbCr, vbNullString))
'If txtValue(Index).Text = vbNullString And cmbField(Index).Text = vbNullString Then Exit Sub
If AddToH(FindH(Index), , True) Then Cancel = True
End Sub

Private Function AddToH(Index As Integer, Optional bolE As Boolean, Optional bolSec As Boolean) As Boolean
Dim ctlF As Control, ctlS As Control, strT(1) As String
If Not bolE Then
Set ctlF = cmbField(Index)
Set ctlS = txtValue(Index)
strT(0) = strHeaders(bytI, Index)
strT(1) = "Additional header"
Else
Set ctlF = txtString(Index)
Set ctlS = txtExp(Index)
strT(0) = strStrings(bytI, Index)
strT(1) = "String"
End If
If strT(0) <> vbNullString Then
If ctlF.Text = Split(strT(0), vbLf)(0) And ctlS.Text = Split(strT(0), vbLf)(1) Then GoTo E1
If ctlF.Text = vbNullString Or ctlS.Text = vbNullString Then
If strT(1) = "String" Then
If Not Filled(bytI, 3) Then
If InStr(cmdOpt(0).Tag, "-" & bytI & "-1" & vbLf) > 0 Then
If txtExp(Index).Tag <> vbNullString Then
If Split(txtExp(Index).Tag, ",")(1) = "1" Or Split(txtExp(Index).Tag, ",")(3) = "1" Then
AddToH = RemI
GoTo E1
End If
End If
End If
End If
End If
If MsgBox("This field will be removed.", vbOKCancel + vbExclamation) = vbOK Then
ctlS.Text = vbNullString
ctlF.Text = vbNullString
If strT(1) = "String" Then
If txtExp(Index).Tag <> vbNullString Then
Dim s() As String: s() = Split(txtExp(Index).Tag, ",")
If s(1) = "1" Then CheckPublic Split(strT(0), vbLf)(0)
If s(1) = "1" Or s(3) = "1" Then StrAR bytI, True, True
End If
txtExp(Index).Tag = vbNullString
cmdOpt(Index).Enabled = False
strStrings(bytI, Index) = vbNullString
cmdAdd(1).Enabled = False
Else
strHeaders(bytI, Index) = vbNullString
cmdAdd(0).Enabled = False
End If
GoTo E
Else: AddToH = True
End If
GoTo E1
End If
ElseIf ctlF.Text = vbNullString Or ctlS.Text = vbNullString Then
If strT(1) <> "String" Then
If Index = cmbField.count - 1 Then cmdAdd(0).Enabled = False
ElseIf Index = txtExp.count - 1 Then cmdAdd(1).Enabled = False
End If
GoTo E1
End If
If strT(1) = "String" Then
If strT(0) <> vbNullString Then
If ctlF.Text <> Split(strT(0), vbLf)(0) Then
CheckPublic Split(strT(0), vbLf)(0)
End If
End If
If strT(0) = vbNullString Or ctlF.Text <> Split(strT(0) & vbLf, vbLf)(0) Then
If txtExp(Index).Tag <> vbNullString Then
If Split(txtExp(Index).Tag, ",")(1) = "1" Then
If CheckPublic(ctlF.Text, True, bolDup) Then
ctlF.SetFocus
ctlF.SelStart = 0
ctlF.SelLength = Len(ctlF.Text)
AddToH = True
GoTo E1
End If
End If
End If
End If
cmdOpt(Index).Enabled = True
strStrings(bytI, Index) = ctlF.Text & vbLf & ctlS.Text & vbLf & txtExp(Index).Tag
If Index = txtExp.count - 1 Then cmdAdd(1).Enabled = True
Else
strHeaders(bytI, Index) = ctlF.Text & vbLf & ctlS.Text
If strT(0) <> vbNullString Then
strT(0) = Left$(strT(0), InStr(strT(0), vbLf) - 1)
If InStr(fraH.Tag, vbLf & " " & strT(0) & " " & vbLf) = 0 Then
If InStr(fraH.Tag, vbLf & " " & ctlF.Text & " " & vbLf) = 0 Then fraH.Tag = vbLf & " " & ctlF.Text & " " & vbLf & ctlS.Text & fraH.Tag Else: fraH.Tag = Replace(fraH.Tag, vbLf & " " & ctlF.Text & " " & vbLf & Split(Split(fraH.Tag, vbLf & " " & ctlF.Text & " " & vbLf)(1), vbLf)(0) & vbLf, vbLf & " " & ctlF.Text & " " & vbLf & ctlS.Text & vbLf)
Else: fraH.Tag = Replace(fraH.Tag, vbLf & " " & strT(0) & " " & vbLf & Split(Split(fraH.Tag, vbLf & " " & strT(0) & " " & vbLf)(1), vbLf)(0) & vbLf, vbLf & " " & ctlF.Text & " " & vbLf & ctlS.Text & vbLf)
End If
Else: If InStr(fraH.Tag, vbLf & " " & ctlF.Text & " " & vbLf) = 0 Then fraH.Tag = vbLf & " " & ctlF.Text & " " & vbLf & ctlS.Text & fraH.Tag Else: fraH.Tag = Replace(fraH.Tag, vbLf & " " & ctlF.Text & " " & vbLf & Split(Split(fraH.Tag, vbLf & " " & ctlF.Text & " " & vbLf)(1), vbLf)(0) & vbLf, vbLf & " " & ctlF.Text & " " & vbLf & ctlS.Text & vbLf)
End If
If Index = cmbField.count - 1 Then cmdAdd(0).Enabled = True
End If
E:
If bolSec And ctlS.Text <> vbNullString Then RplTitle ChkStr(ctlS.Text, 1) Else: RplTitle vbNullString
If bolDebug Then addLog strT(1) & " {index: " & bytI + 1 & ", number: " & Index + 1 & "} -> " & ctlF.Text & ": " & ctlS.Text, True
E1:
Set ctlF = Nothing
Set ctlS = Nothing
End Function

Private Sub txtWait_Validate(Index As Integer, Cancel As Boolean)
If Left$(lblStatus.Caption, 5) = "Shift" Or Left$(lblStatus.Caption, 4) = "Load" Or Left$(lblStatus.Caption, 5) = "Remov" Then Exit Sub
txtWait(Index).Text = Trim$(Replace(Replace(txtWait(Index).Text, vbLf, vbNullString), vbCr, vbNullString))
If Replace(txtWait(Index).Text, "0", vbNullString) = vbNullString And strWait(Index, bytI) = vbNullString Or txtWait(Index).Text = strWait(Index, bytI) Then Exit Sub
If IsNumeric(txtWait(Index).Text) Then txtWait(Index).Text = CLng(txtWait(Index).Text)
strWait(Index, bytI) = txtWait(Index).Text
If txtWait(Index).Text <> vbNullString Then RplTitle ChkStr(txtWait(Index).Text, 1) Else: RplTitle vbNullString
If bolDebug Then addLog "Wait {index: " & bytI + 1 & ", number: " & Index + 1 & "}: " & strWait(Index, bytI), True
End Sub

Private Sub VScroll1_Change(Index As Integer)
PicBox1(Index).Top = -(VScroll1(Index).Value / VScroll1(Index).Max) * (PicBox1(Index).Height - PicBox12(Index).Height)
End Sub

Private Sub VScroll1_Scroll(Index As Integer)
PicBox1(Index).Top = -(VScroll1(Index).Value / VScroll1(Index).Max) * (PicBox1(Index).Height - PicBox12(Index).Height)
End Sub

Private Sub HScroll1_Change()
PicBox2.Left = -(HScroll1.Value / HScroll1.Max) * (PicBox2.Width - PicBox21.Width)
End Sub

Private Sub HScroll1_Scroll()
PicBox2.Left = -(HScroll1.Value / HScroll1.Max) * (PicBox2.Width - PicBox21.Width)
End Sub
