VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmAdminLog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "XOS Plus Administrator - Time Log"
   ClientHeight    =   4350
   ClientLeft      =   645
   ClientTop       =   1305
   ClientWidth     =   6675
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4350
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3480
      Top             =   1080
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About..."
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print Log"
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdViewLog 
      Caption         =   "&View Log"
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdBeginLog 
      Caption         =   "&Begin Log"
      Default         =   -1  'True
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtWorkPlan 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1560
      Width           =   5295
   End
   Begin VB.TextBox txtDate 
      BackColor       =   &H00C0C0C0&
      Height          =   855
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   4095
   End
   Begin ComctlLib.StatusBar lblStatus 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   10
      Top             =   4080
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   476
      Style           =   1
      SimpleText      =   "Ready"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6600
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Please describe your work plan here..."
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   2700
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Administration Utilities"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Plus"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   645
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "XOS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1110
   End
End
Attribute VB_Name = "frmAdminLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()

End Sub

Private Sub cmdBeginLog_Click()
    If Len(txtWorkPlan.Text) < Val(szXOSVar("POLICY_LOGMINCHAR")) Then
        MsgBox "Must enter a " & szXOSVar("POLICY_LOGMINCHAR") & "-character work plan.", 16, "Time Log"
        Exit Sub
    End If
    Open "C:\XOS\USER\WORKLOG" For Append As #1
    lblStatus.SimpleText = "Opened C:\XOS\USER\WORKLOG   " & LOF(1) & " Characters."
    Print #1, "$" & Date$
    Print #1, "User work plan for "; Date$, Time$
    Print #1, txtWorkPlan.Text
    Print #1, "*END"
    
    frmLockDesktop.WindowState = Val(szXOSVar("MISC_DESKTOPSTATE"))
    Unload frmAdminLog
   
End Sub

Private Sub cmdViewLog_Click()
    frmTimelog.Show vbModal
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If szXOSVar("POLICY_ALLOWJUMPOUT") = "TRUE" Then
        If KeyAscii = Val(szXOSVar("KEYSTROKE_JUMPOUT")) Then
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    txtDate.Text = Date$ & "  " & Time$
    Close #1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
 JO = True
End Sub

Private Sub Timer1_Timer()
    txtDate.Text = ProfileRead$("c:\xos\user\config\curuser", "General", "login") & ":   " & Date$ & "  " & Time$
End Sub


Private Sub txtWorkPlan_Change()
    lblStatus.SimpleText = "Work plan:   " & nFigurePercent(Len(txtWorkPlan.Text), Val(szXOSVar("POLICY_LOGMINCHAR"))) & "% of required entry complete.  (" & Len(txtWorkPlan.Text) & " characters)"
    
        
End Sub


Private Sub txtWorkPlan_KeyPress(KeyAscii As Integer)
    If Len(txtWorkPlan.Text) > Val(szXOSVar("POLICY_LOGMAXCHAR")) Then
        txtWorkPlan.Text = Left$(txtWorkPlan.Text, Val(szXOSVar("POLICY_LOGRESETCHAR")))
        Beep
    End If
End Sub


