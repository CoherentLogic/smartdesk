VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmObjectOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Object Options"
   ClientHeight    =   5235
   ClientLeft      =   2025
   ClientTop       =   795
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5235
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save/Exit"
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdDiscard 
      Caption         =   "Discard"
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   4800
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   8070
      _Version        =   393216
      TabOrientation  =   3
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
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
      TabCaption(0)   =   "General"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblIconName"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line1(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Line1(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtTitle"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdChgIcon"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtCommandLine"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Settings"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2(1)"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame3 
         Caption         =   "Font"
         Height          =   1335
         Left            =   -74460
         TabIndex        =   17
         Top             =   2640
         Width           =   4215
         Begin VB.ComboBox cboSize 
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
            Height          =   360
            Left            =   3120
            TabIndex        =   20
            Text            =   "8"
            Top             =   360
            Width           =   855
         End
         Begin VB.ComboBox cboFont 
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
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   240
            Sorted          =   -1  'True
            TabIndex        =   18
            Text            =   "MS Sans Serif"
            Top             =   360
            Width           =   2775
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Colors"
         Height          =   2295
         Left            =   -74460
         TabIndex        =   14
         Top             =   240
         Width           =   4215
         Begin VB.CommandButton cmdChgColor 
            Caption         =   "Change Color..."
            Height          =   375
            Left            =   960
            TabIndex        =   19
            Top             =   840
            Width           =   1335
         End
         Begin VB.ComboBox cboElement 
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
            Height          =   360
            Left            =   960
            Sorted          =   -1  'True
            TabIndex        =   16
            Text            =   "Background Color"
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label lblPreview 
            Appearance      =   0  'Flat
            BackColor       =   &H00808000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Text Preview"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   2400
            TabIndex        =   21
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Element:"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.TextBox txtCommandLine 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1740
         TabIndex        =   13
         Top             =   1680
         Width           =   3375
      End
      Begin VB.Frame Frame1 
         Caption         =   "Current Icon"
         Height          =   1215
         Left            =   540
         TabIndex        =   8
         Top             =   3120
         Width           =   1335
         Begin VB.PictureBox picCurrentIcon 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   360
            ScaleHeight     =   495
            ScaleWidth      =   615
            TabIndex        =   9
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdChgIcon 
         Caption         =   "Change Icon..."
         Height          =   375
         Left            =   1980
         TabIndex        =   3
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox txtTitle 
         BackColor       =   &H00C0FFFF&
         Height          =   855
         Left            =   1140
         TabIndex        =   2
         Text            =   "Default Title"
         Top             =   360
         Width           =   3975
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   540
         X2              =   5460
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   540
         X2              =   5460
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Command Line:"
         Height          =   195
         Left            =   540
         TabIndex        =   12
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblIconName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "None"
         Height          =   195
         Left            =   1980
         TabIndex        =   6
         Top             =   3240
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "XOS Plus"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   1
         Left            =   -70380
         TabIndex        =   5
         Top             =   4200
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "XOS Plus"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   0
         Left            =   4620
         TabIndex        =   4
         Top             =   4200
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Title:"
         Height          =   195
         Left            =   540
         TabIndex        =   1
         Top             =   360
         Width           =   345
      End
   End
End
Attribute VB_Name = "frmObjectOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboFont_Change()
    lblPreview.Font = cboFont.Text
    
End Sub

Private Sub cboFont_Click()
    lblPreview.Font = cboFont.Text
End Sub

Private Sub cboSize_Click()
    lblPreview.FontSize = Val(cboSize.Text)
End Sub

Private Sub cmdChgColor_Click()
    CommonDialog1.ShowColor
    Select Case cboElement.Text
        Case "Text Color"
            CommonDialog1.Color = lblPreview.ForeColor
            lblPreview.ForeColor = CommonDialog1.Color
        Case "Background Color"
            lblPreview.BackColor = CommonDialog1.Color
    End Select
End Sub

Private Sub cmdChgIcon_Click()
    frmIconBrowser.Show
End Sub




Private Sub cmdDiscard_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Call cmdUpdate_Click
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
    Call ProfileWrite("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(CurrentIDX)), "Title", txtTitle.Text)
    Call ProfileWrite("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(CurrentIDX)), "Icon", lblIconName.Caption)
    Call ProfileWrite("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(CurrentIDX)), "Runs", txtCommandLine.Text)
    Call ProfileWrite("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(CurrentIDX)), "Font", cboFont.Text)
    Call ProfileWrite("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(CurrentIDX)), "Size", cboSize.Text)
    'call ProfileWrite("C:\XOS\DESKTOP\DESKMENU.INF",Trim$(str$(CurrentIDX)),"OColorF",
    cmdDiscard.Enabled = False
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    cboElement.AddItem "Background Color"
    cboElement.AddItem "Text Color"
    
    For i = 0 To Printer.FontCount - 1  ' Determine number of fonts.
        cboFont.AddItem Printer.Fonts(i) ' Put each font into list box.

    Next i
    
    For i = 8 To 72 Step 2
        cboSize.AddItem i
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmBlank.Show
    frmMessage.Show
    frmMessage.Refresh
    
    FullQuit = False
    
    Unload frmToolbox
    Unload frmLockDesktop
    frmLockDesktop.Show
    
    Unload frmMessage
    Unload frmBlank
End Sub


Private Sub lblIconName_Change()
    picCurrentIcon.Picture = LoadPicture(lblIconName.Caption)
End Sub

