VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{B16553C3-06DB-101B-85B2-0000C009BE81}#1.0#0"; "SPIN32.OCX"
Begin VB.Form frmRunway 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Runway"
   ClientHeight    =   1380
   ClientLeft      =   180
   ClientTop       =   4680
   ClientWidth     =   6330
   Icon            =   "frmRunway.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1380
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picStationary 
      Height          =   855
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   6075
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.PictureBox picScrolling 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         ScaleHeight     =   585
         ScaleWidth      =   32670
         TabIndex        =   2
         Top             =   120
         Width           =   32700
         Begin Threed.SSCommand cmdProg 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Top             =   120
            Visible         =   0   'False
            Width           =   1125
            _Version        =   65536
            _ExtentX        =   1984
            _ExtentY        =   661
            _StockProps     =   78
            Caption         =   "Empty"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   1
            Font3D          =   3
            RoundedCorners  =   0   'False
            Picture         =   "frmRunway.frx":0442
         End
      End
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   255
      Left            =   5040
      TabIndex        =   6
      Top             =   1080
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Delete"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      Font3D          =   3
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   1080
      Width           =   855
      _Version        =   65536
      _ExtentX        =   1508
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Add..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      Font3D          =   3
   End
   Begin Spin.SpinButton SpinButton1 
      Height          =   255
      Left            =   5880
      TabIndex        =   1
      Top             =   1080
      Width           =   375
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   450
      _StockProps     =   73
      Delay           =   10
      SpinOrientation =   1
      TdThickness     =   1
   End
   Begin VB.Label lblDesc 
      AutoSize        =   -1  'True
      Caption         =   "Ready"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   465
   End
End
Attribute VB_Name = "frmRunway"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdProg_Click(Index As Integer)
    X = Shell(cmdProg(Index).Tag, vbNormalFocus)
End Sub

Private Sub cmdProg_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblDesc.Caption = "Starts " & cmdProg(Index).Caption
End Sub


Private Sub Form_Load()
    FullHeight = frmRunway.Height
    LastIteration = Val(ProfileRead$("\XOS\USER\CONFIG\PROGRAMS.TXT", "Settings", "Programs"))
    
    For Iteration = 0 To LastIteration
        TempItem$ = ProfileRead$("\XOS\USER\CONFIG\PROGRAMS.TXT", Trim(Str(Iteration)), "Description")
        ItemCount = ItemCount + 1
        
        Load cmdProg(ItemCount)
        cmdProg(ItemCount).Tag = ProfileRead$("\XOS\USER\CONFIG\PROGRAMS.TXT", Trim(Str(Iteration)), "ProgramFile")
        cmdProg(ItemCount).Caption = TempItem$
        If Iteration > 0 Then
            If frmRunway.CurrentX > frmRunway.CurrentX + cmdProg(ItemCount).Width + 100 Then
                frmRunway.CurrentX = 100
                frmRunway.CurrentY = frmRunway.CurrentY + cmdProg(ItemCount).Height + 100
                
                FullHeight = FullHeight + cmdProg(ItemCount).Height + 100
                frmRunway.Height = FullHeight
            Else
                frmRunway.CurrentX = frmRunway.CurrentX + cmdProg(ItemCount).Width + 100
            End If
            cmdProg(ItemCount).Left = frmRunway.CurrentX
        Else
            frmRunway.CurrentX = frmRunway.CurrentX + 145
        End If
        cmdProg(ItemCount).Visible = True
        
    Next Iteration
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblDesc = "Ready"
    
    frmRunway.SetFocus
    
End Sub


Private Sub lblDesc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDesc = "Ready"
End Sub


Private Sub picScrolling_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblDesc = "Ready"
End Sub


Private Sub picStationary_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblDesc = "Ready"
End Sub


Private Sub SpinButton1_SpinDown()
    picScrolling.Left = picScrolling.Left + 100
End Sub


Private Sub SpinButton1_SpinUp()
    picScrolling.Left = picScrolling.Left - 100
End Sub


Private Sub Text1_Change()

End Sub


