VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmControlRoom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XOS Control Room"
   ClientHeight    =   5265
   ClientLeft      =   5175
   ClientTop       =   4320
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5265
   ScaleWidth      =   6615
   Begin VB.Timer tmrCRClock 
      Interval        =   1000
      Left            =   4320
      Top             =   4920
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   8281
      _Version        =   393216
      TabOrientation  =   3
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Desktop"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "International"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Date/Time"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Time Log"
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "General"
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame5"
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame5 
         Caption         =   "Fonts"
         Height          =   1095
         Left            =   -74760
         TabIndex        =   13
         Top             =   3240
         Width           =   4095
         Begin VB.CommandButton Command1 
            Caption         =   "Font Browser..."
            Height          =   495
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Date"
         Height          =   3135
         Left            =   -74880
         TabIndex        =   11
         Top             =   240
         Width           =   5655
         Begin MSACAL.Calendar Calendar1 
            Height          =   2415
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   5175
            _Version        =   524288
            _ExtentX        =   9128
            _ExtentY        =   4260
            _StockProps     =   1
            BackColor       =   -2147483633
            Year            =   1996
            Month           =   3
            Day             =   14
            DayLength       =   1
            MonthLength     =   2
            DayFontColor    =   0
            FirstDay        =   1
            GridCellEffect  =   2
            GridFontColor   =   10485760
            GridLinesColor  =   -2147483632
            ShowDateSelectors=   -1  'True
            ShowDays        =   -1  'True
            ShowHorizontalGrid=   -1  'True
            ShowTitle       =   -1  'True
            ShowVerticalGrid=   -1  'True
            TitleFontColor  =   10485760
            ValueIsNull     =   0   'False
            BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Desktop &Wallpaper"
         Height          =   2415
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   5535
         Begin VB.Frame Frame3 
            Caption         =   "Show &Paper"
            Height          =   1815
            Left            =   3480
            TabIndex        =   7
            Top             =   360
            Width           =   1815
            Begin VB.OptionButton ShowPaper 
               Caption         =   "Not at all"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   10
               Top             =   1080
               Width           =   1455
            End
            Begin VB.OptionButton ShowPaper 
               Caption         =   "Menu Option"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   9
               Top             =   720
               Width           =   1455
            End
            Begin VB.OptionButton ShowPaper 
               Caption         =   "Automatic"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   8
               Top             =   360
               Value           =   -1  'True
               Width           =   1455
            End
         End
         Begin VB.Image imgPreview 
            Height          =   1455
            Left            =   240
            Stretch         =   -1  'True
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Preview:"
            Height          =   195
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "&Show Desktop"
         Height          =   1455
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   2415
         Begin VB.OptionButton OptState 
            Caption         =   "Normal"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   1
            Top             =   360
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton OptState 
            Caption         =   "Maximized"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   4
            Top             =   1080
            Width           =   1815
         End
         Begin VB.OptionButton OptState 
            Caption         =   "Minimized"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   3
            Top             =   720
            Width           =   1815
         End
      End
   End
End
Attribute VB_Name = "frmControlRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmFontBrowser.Show

End Sub


Private Sub Form_Load()
    
    Select Case szXOSVar("MISC_DESKTOPSTATE")
        Case "0"
            OptState(0).Value = True
        Case "1"
            OptState(1).Value = True
        Case "2"
            OptState(3).Value = True
    End Select
    imgPreview.Picture = LoadPicture(szXOSVar("MISC_DESKTOPPICTURE"))
    
    Select Case szXOSVar("MISC_LOADPICTURE")
        Case "TRUE"
            ShowPaper(0).Value = True
        Case "FALSE"
            ShowPaper(1).Value = True
    End Select
    
    If szXOSVar("MISC_LOADPICMENU") = "FALSE" Then
        ShowPaper(2).Value = True
    End If
    
End Sub

Private Sub Option2_Click()

End Sub


