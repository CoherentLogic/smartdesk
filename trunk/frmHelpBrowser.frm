VERSION 5.00
Begin VB.Form frmHelpBrowser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help Browser"
   ClientHeight    =   3645
   ClientLeft      =   2610
   ClientTop       =   1995
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3645
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   3240
      Width           =   1095
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00C0FFFF&
      Height          =   2820
      Left            =   120
      Pattern         =   "*.xmf"
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmHelpBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    HelpFile = File1.FileName
    frmHelpQuery.txtHFile.Text = HelpFile
    Unload Me
End Sub

Private Sub Form_Load()
    File1.Path = szXOSVar("DEFAULT_XMF")
End Sub


