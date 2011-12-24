VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMessage 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1530
   ClientLeft      =   2745
   ClientTop       =   2235
   ClientWidth     =   4275
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1530
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Please Wait"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "XOS Plus"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   0
      Left            =   3240
      TabIndex        =   1
      Top             =   1200
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Updating System Configuration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   4335
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Refresh
    Call MakeTopWindow(frmMessage)
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    ProgressBar1.Value = 0
End Sub


Private Sub Form_LostFocus()
    Call MakeTopWindow(frmMessage)
End Sub

