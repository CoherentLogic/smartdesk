VERSION 5.00
Begin VB.Form frmStopwatch 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stopwatch"
   ClientHeight    =   1635
   ClientLeft      =   4050
   ClientTop       =   2550
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1635
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkAlarm 
      Caption         =   "Alarm"
      Height          =   255
      Left            =   3720
      TabIndex        =   13
      Top             =   1200
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CommandButton cmdSetTimer 
      Caption         =   "Set Timer"
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdResume 
      Caption         =   "Resume"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   1800
      Top             =   2400
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1200
      Top             =   2400
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   720
      Top             =   2400
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   2400
   End
   Begin VB.Line Line1 
      X1              =   2400
      X2              =   2400
      Y1              =   120
      Y2              =   1560
   End
   Begin VB.Label lblTimer 
      AutoSize        =   -1  'True
      Caption         =   "Not Set"
      Height          =   195
      Left            =   1320
      TabIndex        =   11
      Top             =   1200
      Width           =   540
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "Timer:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   435
   End
   Begin VB.Label lblSince 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1320
      TabIndex        =   9
      Top             =   840
      Width           =   45
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Since Midnight:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblCTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10:30:25 AM"
      Height          =   195
      Left            =   1320
      TabIndex        =   7
      Top             =   480
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Current Time:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Stopwatch:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "00:00.00"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   645
   End
End
Attribute VB_Name = "frmStopwatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdPause_Click()
    Timer3.Enabled = False
    cmdResume.Enabled = True
    cmdPause.Enabled = False
End Sub


Private Sub cmdResume_Click()
    Timer3.Enabled = True
    cmdResume.Enabled = False
    cmdPause.Enabled = True
End Sub

Private Sub cmdStart_Click()
    Minutes = 0
    Seconds = 0
    Tenths = 0
    cmdStop.Enabled = True
    cmdPause.Enabled = True
    cmdStart.Enabled = False
    
    Timer3.Enabled = True
End Sub

Private Sub cmdStop_Click()
    Timer3.Enabled = False
    MsgBox "Elapsed Time:" & vbCrLf & vbCrLf & Label1.Caption, vbInformation, "Stopwatch Applet"
    Label1 = "00:00.00"
    cmdStart.Enabled = True
    cmdPause.Enabled = False
    cmdResume.Enabled = False
    cmdStop.Enabled = False
End Sub

Private Sub Form_Load()
    frmToolbox.imgStopwatch.Visible = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmToolbox.imgStopwatch.Visible = False
End Sub


Private Sub Timer1_Timer()
    Tenths = Tenths + 1
    lblSince = Timer
    If Tenths >= 10 Then
        Tenths = 0
    End If
End Sub


Private Sub Timer2_Timer()
    Seconds = Seconds + 1
    lblCTime = Format$(Now, "h:mm:ss AM/PM")
    
    If Seconds >= 60 Then
        Seconds = 0
        Minutes = Minutes + 1
    End If
End Sub


Private Sub Timer3_Timer()
    Label1.Caption = Format$(Minutes, "##00") & ":" & Format$(Seconds, "##00") & "." & Format$(Tenths, "##00")
End Sub


Private Sub Timer4_Timer()
    Hundredths = Hundredths + 1
    If Hundredths >= 100 Then
        Hundredths = 0
    End If
End Sub


