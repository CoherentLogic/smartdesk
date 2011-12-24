VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmRefAsst 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Reference Assistant"
   ClientHeight    =   4095
   ClientLeft      =   645
   ClientTop       =   1635
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4095
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picSet1 
      BorderStyle     =   0  'None
      Height          =   2055
      Index           =   0
      Left            =   3600
      ScaleHeight     =   2055
      ScaleWidth      =   4335
      TabIndex        =   3
      Top             =   1200
      Width           =   4335
      Begin VB.TextBox txtAppTitle 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   15
         Top             =   1320
         Width           =   3375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Browse..."
         Height          =   375
         Left            =   3480
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtCommandLine 
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
         Left            =   0
         TabIndex        =   5
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Please enter the title of the application:"
         Height          =   195
         Left            =   0
         TabIndex        =   14
         Top             =   1080
         Width           =   2745
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Please enter the location of the program you wish to add to the desktop:"
         Height          =   390
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   3180
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404080&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   3600
      ScaleHeight     =   855
      ScaleWidth      =   4335
      TabIndex        =   1
      Top             =   120
      Width           =   4335
      Begin VB.Label lblStep 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step 1 of 2 - Select Program"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   2280
      End
      Begin VB.Image Image3 
         Height          =   375
         Left            =   3840
         Stretch         =   -1  'True
         Top             =   240
         Width           =   375
      End
      Begin VB.Image Image2 
         Height          =   375
         Left            =   3480
         Stretch         =   -1  'True
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Reference"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   2640
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      _Version        =   65536
      _ExtentX        =   5953
      _ExtentY        =   6800
      _StockProps     =   15
      Caption         =   "SSPanel1"
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Autosize        =   3
      Begin VB.Image Image1 
         Height          =   3705
         Left            =   75
         Stretch         =   -1  'True
         Top             =   75
         Width           =   3225
      End
   End
   Begin VB.PictureBox picSet1 
      BorderStyle     =   0  'None
      Height          =   2535
      Index           =   1
      Left            =   3600
      ScaleHeight     =   2535
      ScaleWidth      =   4335
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   4335
      Begin VB.Frame Frame1 
         Caption         =   "Icon"
         Height          =   1215
         Left            =   2520
         TabIndex        =   12
         Top             =   480
         Width           =   1215
         Begin VB.PictureBox picIcon 
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   360
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   13
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Click Here To Select"
         Height          =   375
         Left            =   2160
         TabIndex        =   11
         Top             =   2160
         Width           =   2175
      End
      Begin VB.FileListBox filIcon 
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
         Height          =   1530
         Left            =   0
         Pattern         =   "*.ico"
         TabIndex        =   9
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Please select an icon to represent the program on the desktop:"
         Height          =   495
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   2775
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3000
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRefAsst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    CommonDialog1.Filter = szXOSVar("MISC_DEFAULTFILTER")
    'CommonDialog1.DialogCaption = "Select Program"
    CommonDialog1.ShowOpen
    txtCommandLine.Text = CommonDialog1.FileName
End Sub

Private Sub Command2_Click()
    Call ProfileWrite("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(CurrentIDX)), "Runs", txtCommandLine.Text)
    Call ProfileWrite("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(CurrentIDX)), "Title", txtAppTitle.Text)
    Call ProfileWrite("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(CurrentIDX)), "Icon", "C:\XOS\DESKTOP\ICONS\" & filIcon.FileName)
    frmLockDesktop.picDTIcon(CurrentIDX).Tag = txtCommandLine.Text
    frmLockDesktop.picDTIcon(CurrentIDX).Picture = LoadPicture("C:\XOS\DESKTOP\ICONS\" & filIcon.FileName)
    frmLockDesktop.picDTIcon(CurrentIDX).CurrentY = 500
    frmLockDesktop.picDTIcon(CurrentIDX).Width = frmLockDesktop.picDTIcon(i).TextWidth(ProfileRead$("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(CurrentIDX)), "Title"))
    frmLockDesktop.picDTIcon(CurrentIDX).Print ProfileRead$("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(CurrentIDX)), "Title")
    Unload Me
End Sub


Private Sub filIcon_Click()
    picIcon.Picture = LoadPicture("C:\XOS\DESKTOP\ICONS\" & filIcon.FileName)
End Sub

Private Sub Form_Load()
    filIcon.Path = "C:\XOS\DESKTOP\ICONS"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ActiveFrame = 0
End Sub


Private Sub Image2_Click()
    On Error Resume Next
    If Not ActiveFrame = 0 Then
        ActiveFrame = ActiveFrame - 1
        
        picSet1(ActiveFrame + 1).Visible = False
        picSet1(ActiveFrame).Visible = True
        Select Case ActiveFrame
            Case 0
                lblStep = "Step 1 of 2 - Select Program"
            Case 1
                lblStep = "Step 2 of 2 - Select Icon"
        End Select
    End If
End Sub

Private Sub Image3_Click()
    If Not ActiveFrame = 1 Then
        ActiveFrame = ActiveFrame + 1
        picSet1(ActiveFrame - 1).Visible = False
        picSet1(ActiveFrame).Visible = True
        Select Case ActiveFrame
            Case 0
                lblStep = "Step 1 of 2 - Select Program"
            Case 1
                lblStep = "Step 2 of 2 - Select Icon"
        End Select
    End If
End Sub


