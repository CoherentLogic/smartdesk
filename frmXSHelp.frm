VERSION 5.00
Begin VB.Form frmXSHelp 
   Caption         =   "XOS Plus - Help"
   ClientHeight    =   4200
   ClientLeft      =   3945
   ClientTop       =   2655
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4200
   ScaleWidth      =   5520
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   6735
      TabIndex        =   1
      Top             =   0
      Width           =   6735
      Begin VB.CommandButton Command4 
         Caption         =   "&Print"
         Height          =   375
         Left            =   2520
         TabIndex        =   5
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   3960
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&History..."
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Query..."
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.TextBox txtHelpText 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   600
      Width           =   6735
   End
End
Attribute VB_Name = "frmXSHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    frmHelpQuery.Show
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    On Error Resume Next
    Printer.Print txtHelpText.Text
    Printer.EndDoc
End Sub

Private Sub Form_Resize()
    picToolBar.Width = frmXSHelp.Width
    txtHelpText.Width = frmXSHelp.ScaleWidth
    txtHelpText.Height = frmXSHelp.ScaleHeight - picToolBar.Height
End Sub


