VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmXS9PIntro 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4140
   ClientLeft      =   1185
   ClientTop       =   1845
   ClientWidth     =   6690
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4140
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1095
      ScaleWidth      =   3135
      TabIndex        =   2
      Top             =   2760
      Width           =   3135
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R E L E A S E   X"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plus"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   825
         Left            =   1440
         TabIndex        =   4
         Top             =   0
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "XOS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1455
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   2895
      Left            =   6600
      TabIndex        =   0
      Top             =   3960
      Visible         =   0   'False
      Width           =   6495
      _Version        =   65536
      _ExtentX        =   11456
      _ExtentY        =   5106
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   1
      BevelInner      =   2
      Begin VB.Timer Timer1 
         Interval        =   5000
         Left            =   4200
         Top             =   2160
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "97"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   645
      Left            =   2640
      TabIndex        =   8
      Top             =   360
      Width           =   510
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SmartDesk"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   2430
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "Label6"
      Height          =   1095
      Left            =   360
      TabIndex        =   6
      Top             =   2880
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Internet Enabled Desktop"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   5775
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5880
      Top             =   360
      Width           =   480
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6600
      Y1              =   960
      Y2              =   960
   End
End
Attribute VB_Name = "frmXS9PIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'    Me.Left = (Screen.Width - Me.Width) / 2
 '   Me.Top = (Screen.Height - Me.Height) / 2
End Sub

Private Sub Timer1_Timer()
    Unload Me
    Unload frmBlank
    Timer1.Enabled = False
End Sub


