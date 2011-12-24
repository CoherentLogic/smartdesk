VERSION 5.00
Begin VB.Form dlgGetURL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open Internet Site"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Go"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtURL 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "WWW Address:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1155
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4440
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Browse the Internet"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2130
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4440
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "dlgGetURL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGo_Click()
    StartingAddress = txtURL.Text
    frmBrowser.Show
    frmBrowser.WindowState = 2
    Unload Me
End Sub
