VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmToolbox 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   870
   ClientLeft      =   60
   ClientTop       =   6045
   ClientWidth     =   15240
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   870
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   WhatsThisHelp   =   -1  'True
   Begin Threed.SSCommand cmdReload 
      Height          =   375
      Left            =   7920
      TabIndex        =   10
      ToolTipText     =   "Reload the desktop"
      Top             =   0
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Reload"
      Font3D          =   3
   End
   Begin Threed.SSCommand cmdXOSShell 
      Height          =   375
      Left            =   6600
      TabIndex        =   9
      ToolTipText     =   "Display the XOS Shell prompt"
      Top             =   480
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "XOS Shell"
      Font3D          =   3
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   375
      Left            =   7920
      TabIndex        =   8
      ToolTipText     =   "Exit XOS and return to Windows Explorer"
      Top             =   480
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Exit"
      Font3D          =   3
   End
   Begin Threed.SSCommand cmdShutdown 
      Height          =   375
      Left            =   6600
      TabIndex        =   7
      ToolTipText     =   "Shut Down the Computer"
      Top             =   0
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Shut Down"
      Font3D          =   3
   End
   Begin Threed.SSCommand cmdHideToolbox 
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      ToolTipText     =   "Hide this toolbar"
      Top             =   480
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Hide Toolbar"
      Font3D          =   3
   End
   Begin Threed.SSCommand cmdClassicMenu 
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      ToolTipText     =   "Display Classic Menu"
      Top             =   0
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Classic Menu"
      Font3D          =   3
   End
   Begin VB.PictureBox picRight 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   9960
      ScaleHeight     =   735
      ScaleWidth      =   5175
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.CommandButton Command1 
         Caption         =   "Run"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   3
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   1935
      End
      Begin Threed.SSPanel pnlClock 
         Height          =   495
         Left            =   3000
         TabIndex        =   4
         Top             =   120
         Width           =   2055
         _Version        =   65536
         _ExtentX        =   3625
         _ExtentY        =   873
         _StockProps     =   15
         Caption         =   "0:00 PM"
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
         Font3D          =   3
         Alignment       =   4
         Begin VB.Image imgStopwatch 
            Height          =   345
            Left            =   720
            Stretch         =   -1  'True
            Top             =   120
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Image imgDiskView 
            Height          =   315
            Left            =   360
            Stretch         =   -1  'True
            Top             =   120
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Image imgNetConnect 
            Height          =   315
            Left            =   0
            OLEDropMode     =   1  'Manual
            Stretch         =   -1  'True
            Top             =   120
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "XOS Plus SmartDesk Toolbar"
         ForeColor       =   &H00808080&
         Height          =   210
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   2610
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7920
      Top             =   600
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   4200
      ToolTipText     =   "View Drives Object"
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   3480
      ToolTipText     =   "Open Internet Site"
      Top             =   120
      Width           =   480
   End
   Begin VB.Line Line1 
      X1              =   5040
      X2              =   5040
      Y1              =   0
      Y2              =   840
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   2760
      ToolTipText     =   "New Help Query"
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   2040
      ToolTipText     =   "XOS Control Room"
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   1320
      ToolTipText     =   "Profile Editor"
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   720
      ToolTipText     =   "Hide Desktop"
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      ToolTipText     =   "Desktop View"
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmToolbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClassicMenu_Click()
    Select Case frmLockDesktop.picRunMenu.Visible
        Case True
            frmLockDesktop.picRunMenu.Visible = False
        Case False
            frmLockDesktop.picRunMenu.Visible = True
            frmLockDesktop.picRunMenu.Top = Screen.Height - frmLockDesktop.picRunMenu.Height - frmToolbox.Height - 250
            frmLockDesktop.picRunMenu.Left = cmdClassicMenu.Left
            
    End Select
End Sub

Private Sub cmdExit_Click()
    FullQuit = True
    Unload frmLockDesktop
End Sub


Private Sub cmdHideToolbox_Click()
    frmLockDesktop.mnuToolbar.Checked = False
    Me.Hide
End Sub

Private Sub cmdReload_Click()
    Call ReloadDesktop
End Sub

Private Sub cmdShutdown_Click()
        If szXOSVar("MISC_CONFIRM") = True Then
        
        End If
        frmMessage.Label1(0) = "System Shutdown"
        frmMessage.Label1(1) = "Please Wait"
        frmMessage.Show
        frmMessage.Refresh
        Delay 2
        X = ExitWindowsEx(EWX_SHUTDOWN, 0)
        End
End Sub


Private Sub cmdXOSShell_Click()
dummy = Shell("C:\XOS\SYSTEM\exSH.EXE", vbNormalFocus)

End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim WinText As String * 100
Dim TaskHWnd As Long
Dim boom As Long
If Mid$(Text1.Text, 4, 1) = ":" Then
    Select Case Left$(Text1.Text, 3)
        Case "int"
        '
        'TODO: add code to load internal objects here
        '
        Case "dvw"
            CDir = Mid$(Text1.Text, 5)
            frmDiskViewer.Show
            Exit Sub
    End Select
End If

X = Shell(Text1.Text, vbMinimizedFocus)
RunIndex = RunIndex + 1
RunList(RunIndex) = X
AppActivate X
TaskHWnd = GetForegroundWindow()

boom = GetWindowText(TaskHWnd, WinText, 100)
RunListText(RunIndex) = Trim$(WinText)
End Sub

Private Sub Form_Load()
    Call MakeTopWindow(frmToolbox)
    frmToolbox.Width = Screen.Width
    frmToolbox.Top = Screen.Height - Me.Height
    picRight.Left = Screen.Width - picRight.Width
End Sub


Private Sub Form_Resize()
    Call MakeTopWindow(frmToolbox)
End Sub


Private Sub Image1_Click()
    frmLockDesktop.Show
    frmLockDesktop.WindowState = 2
    
End Sub


Private Sub Image2_Click()
    frmLockDesktop.WindowState = 1
End Sub


Private Sub Image3_Click()
    frmShowVar.Show
End Sub


Private Sub Image4_Click()
    frmControlRoom.Show
End Sub


Private Sub Image5_Click()
    frmHelpQuery.Show
End Sub


Private Sub Image6_Click()
    frmSelectDrive.Show
End Sub
Private Sub Image7_Click()
    dlgGetURL.Show
End Sub

Private Sub imgDiskView_Click()
    frmDiskViewer.SetFocus
End Sub

Private Sub imgNetConnect_Click()
    frmBrowser.SetFocus
End Sub

Private Sub imgStopwatch_Click()
    frmStopwatch.SetFocus
End Sub

Private Sub Timer1_Timer()
    pnlClock.Caption = Format$(Now, szXOSVar("LOCALE_TIMEFORMAT"))
    pnlClock.ToolTipText = Format(Date, "Long Date")


End Sub
