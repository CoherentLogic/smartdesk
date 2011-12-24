VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmDiskViewer 
   Caption         =   "Disk Viewer"
   ClientHeight    =   8865
   ClientLeft      =   2085
   ClientTop       =   1635
   ClientWidth     =   11580
   LinkTopic       =   "Form2"
   ScaleHeight     =   8865
   ScaleWidth      =   11580
   Begin Threed.SSPanel pnlFiles 
      Height          =   5775
      Left            =   3960
      TabIndex        =   8
      Top             =   600
      Width           =   7575
      _Version        =   65536
      _ExtentX        =   13361
      _ExtentY        =   10186
      _StockProps     =   15
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
      BevelOuter      =   1
      Begin VB.FileListBox File1 
         BackColor       =   &H00C0FFFF&
         Height          =   5550
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   7335
      End
   End
   Begin Threed.SSPanel pnlProperties 
      Height          =   7695
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   3855
      _Version        =   65536
      _ExtentX        =   6800
      _ExtentY        =   13573
      _StockProps     =   15
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
      BevelOuter      =   1
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open..."
         Height          =   375
         Left            =   2160
         TabIndex        =   16
         Top             =   7200
         Width           =   1455
      End
      Begin VB.TextBox txtPreview 
         Height          =   2295
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   4800
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox txtPattern 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1800
         TabIndex        =   12
         Text            =   "*.*"
         Top             =   3480
         Width           =   1815
      End
      Begin VB.CheckBox chkArchive 
         Caption         =   "Archive"
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   3000
         Width           =   1455
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "System"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   3000
         Width           =   975
      End
      Begin VB.CheckBox chkReadOnly 
         Caption         =   "Read-Only"
         Height          =   255
         Left            =   1920
         TabIndex        =   6
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CheckBox chkHidden 
         Caption         =   "Hidden"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2520
         Width           =   975
      End
      Begin VB.DirListBox Dir1 
         BackColor       =   &H00C0FFFF&
         Height          =   1440
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   3615
      End
      Begin VB.Image imgPicPreview 
         Height          =   2295
         Left            =   240
         Stretch         =   -1  'True
         Top             =   4800
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   3720
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "File Preview:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   4320
         Width           =   900
      End
      Begin VB.Label lblFileType 
         Alignment       =   2  'Center
         Caption         =   "File Type"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   3960
         Width           =   3375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Show Files of Type:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   3480
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "File Properties:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Current Folder:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1035
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   3720
         Y1              =   2280
         Y2              =   2280
      End
   End
   Begin VB.PictureBox picToolBar 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   11580
      TabIndex        =   0
      Top             =   0
      Width           =   11580
      Begin VB.Image Image7 
         Height          =   240
         Left            =   480
         Top             =   120
         Width           =   240
      End
      Begin VB.Line Line4 
         X1              =   2760
         X2              =   2760
         Y1              =   0
         Y2              =   480
      End
      Begin VB.Image Image6 
         Height          =   240
         Left            =   2400
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Image5 
         Height          =   240
         Left            =   2040
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Image4 
         Height          =   240
         Left            =   1680
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   1320
         Top             =   120
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   960
         Top             =   120
         Width           =   240
      End
      Begin VB.Line Line3 
         X1              =   840
         X2              =   840
         Y1              =   0
         Y2              =   480
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   120
         Top             =   120
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmDiskViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    frmDiskViewer.Caption = "Disk Viewer - " & Dir1.Path
    frmToolbox.imgDiskView.ToolTipText = frmDiskViewer.Caption
    
    
End Sub

Private Sub File1_Click()
    On Error Resume Next
    frmDiskViewer.Caption = "Disk Viewer - " & Dir1.Path
    lblFileType.Caption = ProfileRead$("C:\XOS\XOS.XIF", "Extensions", LCase$(Right$(File1.FileName, 3)))
    Previewer$ = ProfileRead$("C:\XOS\XOS.XIF", "Viewers", LCase$(Right$(File1.FileName, 3)))
    imgPicPreview.Visible = False
    txtPreview.Visible = False
    
    If Left$(Previewer$, 4) = "int:" Then
        Select Case Mid$(Previewer$, 5)
            Case "picture"
                imgPicPreview.Picture = LoadPicture(szGetRootDir(Dir1.Path, File1.FileName))
                imgPicPreview.Visible = True
            Case "text"
                dummy = FreeFile
                Open szGetRootDir(Dir1.Path, File1.FileName) For Input As #dummy
                txtPreview.Text = Input$(LOF(dummy), dummy)
                Close #dummy
                txtPreview.Visible = True
                
        End Select
    Else
    '
    'TODO: Add external viewer code here.
    '
    End If
    
    Result = GetAttr(szGetRootDir(Dir1.Path, File1.FileName)) And vbArchive
    If Result <> 0 Then
        chkArchive.Value = vbChecked
    Else
        chkArchive.Value = vbUnchecked
    End If
    Result = GetAttr(szGetRootDir(Dir1.Path, File1.FileName)) And vbSystem
    If Result <> 0 Then
        chkSystem.Value = vbChecked
    Else
        chkSystem.Value = vbUnchecked
    End If
    Result = GetAttr(szGetRootDir(Dir1.Path, File1.FileName)) And vbHidden
    If Result <> 0 Then
        chkHidden.Value = vbChecked
    Else
        chkHidden.Value = vbUnchecked
    End If
    Result = GetAttr(szGetRootDir(Dir1.Path, File1.FileName)) And vbReadOnly
    If Result <> 0 Then
        chkReadOnly.Value = vbChecked
    Else
        chkReadOnly.Value = vbUnchecked
    End If
    
    frmToolbox.imgDiskView.ToolTipText = frmDiskViewer.Caption
        
    
End Sub

Private Sub Form_Load()
    Dir1.Path = CDir
    frmToolbox.imgDiskView.ToolTipText = "Disk Viewer"
    frmToolbox.imgDiskView.Visible = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmToolbox.imgDiskView.Visible = False
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    pnlProperties.Top = picToolBar.Height + 10
    pnlProperties.Height = Me.ScaleHeight - picToolBar.Height - 10
    pnlFiles.Height = Me.ScaleHeight - picToolBar.Height - 10
    pnlFiles.Top = picToolBar.Height + 10
    pnlFiles.Width = Me.ScaleWidth - pnlProperties.Width - 150
    File1.Height = pnlFiles.Height - 150
    File1.Width = pnlFiles.Width - 150
    
End Sub


Private Sub txtPattern_Change()
    On Error Resume Next
    File1.Pattern = txtPattern.Text
    
End Sub


