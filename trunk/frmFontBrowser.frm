VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmFontBrowser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Font Browser"
   ClientHeight    =   6525
   ClientLeft      =   4590
   ClientTop       =   1545
   ClientWidth     =   4950
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   Begin Threed.SSPanel SSPanel1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      _Version        =   65536
      _ExtentX        =   8281
      _ExtentY        =   9763
      _StockProps     =   15
      BackColor       =   11059392
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
      Begin VB.PictureBox picFont 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   32000
         Left            =   0
         ScaleHeight     =   31995
         ScaleWidth      =   4665
         TabIndex        =   1
         Top             =   -120
         Width           =   4665
      End
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   240
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   720
      Top             =   5880
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   1800
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   1800
      Top             =   3240
      Width           =   1215
   End
End
Attribute VB_Name = "frmFontBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    frmMessage.Show
    frmMessage.Label1(0).Caption = "Loading Fonts..."
    frmMessage.Label1(1).Caption = "Please Wait"
    frmMessage.ProgressBar1.Max = Printer.FontCount - 1
    For i = 1 To Printer.FontCount - 1
        picFont.FontName = Printer.Fonts(i)
        
        picFont.Print Printer.Fonts(i)
        frmMessage.ProgressBar1.Value = i
        frmMessage.Label1(0).Caption = "Loading Font " & Printer.Fonts(i)
        frmMessage.Label1(0).Refresh
        
    Next
        Unload frmMessage
End Sub


Private Sub Image3_Click()
picFont.Top = picFont.Top - 1

End Sub


Private Sub Image4_Click()
picFont.Top = picFont.Top + 1

End Sub


