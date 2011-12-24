VERSION 5.00
Begin VB.Form frmIconBrowser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Icon Browser"
   ClientHeight    =   3960
   ClientLeft      =   3435
   ClientTop       =   1950
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3960
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   3000
      Width           =   1335
   End
   Begin VB.PictureBox picPreview 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   4680
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2970
      Left            =   120
      Pattern         =   "*.ico"
      TabIndex        =   0
      Top             =   480
      Width           =   3495
   End
   Begin VB.Line Line1 
      X1              =   3840
      X2              =   3840
      Y1              =   0
      Y2              =   3960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Icon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3960
      TabIndex        =   3
      Top             =   120
      Width           =   390
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Left            =   4560
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Icon Files:"
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmIconBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    NewIcon = File1.Path & "\" & File1.FileName
    frmObjectOptions.lblIconName = NewIcon
    Unload Me
End Sub


Private Sub File1_Click()
    picPreview.Picture = LoadPicture(File1.Path & "\" & File1.FileName)
End Sub


Private Sub Form_Load()
    File1.Path = "C:\XOS\DESKTOP\ICONS"
End Sub


Private Sub Label2_Click()

End Sub


