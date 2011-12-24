VERSION 5.00
Begin VB.Form frmTimelog 
   Caption         =   "Time Log Viewer"
   ClientHeight    =   4680
   ClientLeft      =   765
   ClientTop       =   1050
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4680
   ScaleWidth      =   6750
   Begin VB.CommandButton cmdShowEndLog 
      Caption         =   "Show &End Log"
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdShowLog 
      Caption         =   "&Show Start Log"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtViewLog 
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
      Height          =   3495
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   600
      Width           =   6495
   End
   Begin VB.TextBox txtDate 
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
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter a date to view here:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2235
   End
End
Attribute VB_Name = "frmTimelog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdShowEndLog_Click()
    
    txtViewLog.Text = szExtractDataSection(Temp$, txtDate.Text & "-END")
    
End Sub

Private Sub cmdShowLog_Click()
    
    txtViewLog.Text = szExtractDataSection(Temp$, txtDate.Text)
    
    
End Sub

Private Sub Form_Load()
    Close #1
    txtDate = Date$
    Temp$ = szReadFromFile("C:\XOS\USER\WORKLOG")
    Call CloseFile
    Open "C:\XOS\USER\WORKLOG" For Output As #1
End Sub


