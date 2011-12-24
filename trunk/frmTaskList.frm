VERSION 5.00
Begin VB.Form frmTaskList 
   Caption         =   "Task List"
   ClientHeight    =   2775
   ClientLeft      =   2445
   ClientTop       =   1545
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2775
   ScaleWidth      =   5655
   Begin VB.PictureBox picToolbar 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   4095
      ScaleHeight     =   2775
      ScaleWidth      =   1560
      TabIndex        =   1
      Top             =   0
      Width           =   1560
      Begin VB.CommandButton Command3 
         Caption         =   "&End Task"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Switch To..."
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.ListBox lstTasks 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "frmTaskList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Private Sub Command1_Click()
    On Error Resume Next
    AppActivate Str$(lstTasks.Text)
    TheApp$ = lstTasks.Text
    SendKeys "% R"
    Unload Me
    AppActivate TheApp$
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    AppActivate lstTasks.Text
    TheApp$ = lstTasks.Text
    SendKeys "% C"
    Unload Me
    AppActivate TheApp$
End Sub


Private Sub Form_Load()
    For i = 1 To RunIndex
        lstTasks.AddItem RunListText(i)
    Next
End Sub

Private Sub Form_Resize()
    lstTasks.Width = Me.ScaleWidth - picToolBar.Width
    lstTasks.Height = Me.ScaleHeight
    
End Sub


Private Sub lstTasks_DblClick()
    
    
    AppActivate lstTasks.Text
    TheApp$ = lstTasks.Text
    SendKeys "% R"
    Unload Me
    AppActivate TheApp$

End Sub


