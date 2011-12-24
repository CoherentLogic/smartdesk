VERSION 5.00
Begin VB.Form frmBlank 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4140
   ClientLeft      =   1875
   ClientTop       =   2580
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4140
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3120
      Top             =   2520
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Desktop"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4200
      Top             =   2880
      Width           =   480
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   4320
      TabIndex        =   1
      Top             =   5880
      Width           =   615
   End
End
Attribute VB_Name = "frmBlank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReturn_Click()
    Unload Me
    cmdReturn.Visible = False
    frmToolbox.Show
End Sub

Private Sub Form_Load()
    'Picture = LoadPicture(szXOSVar("MISC_DESKTOPPICTURE"))
    
End Sub


Private Sub Timer1_Timer()
    On Error Resume Next
    Image1.Left = Rnd * frmBlank.Width
    Image1.Top = Rnd * frmBlank.Height
    
End Sub


