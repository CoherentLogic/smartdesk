VERSION 5.00
Begin VB.Form frmSelectDrive 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Disk Viewer - Select Drive"
   ClientHeight    =   3645
   ClientLeft      =   8775
   ClientTop       =   5805
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3645
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.DirListBox dirStartDir 
      BackColor       =   &H00C0FFFF&
      Height          =   1440
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   3975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   3975
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   4200
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Drive"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1740
   End
End
Attribute VB_Name = "frmSelectDrive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub


Private Sub cmdOK_Click()
    CDrive = Drive1.Drive
    CDir = dirStartDir.Path
    frmDiskViewer.Show
    Unload Me
End Sub


Private Sub Drive1_Change()
    dirStartDir.Path = Drive1.Drive
End Sub


