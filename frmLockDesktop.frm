VERSION 5.00
Begin VB.Form frmLockDesktop 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   ClientHeight    =   8370
   ClientLeft      =   1560
   ClientTop       =   1605
   ClientWidth     =   15420
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8370
   ScaleWidth      =   15420
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrTaskUpdate 
      Interval        =   5000
      Left            =   10560
      Top             =   5760
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   9255
      Left            =   12600
      ScaleHeight     =   9225
      ScaleWidth      =   2625
      TabIndex        =   10
      Top             =   240
      Width           =   2655
      Begin VB.CommandButton cmdTask 
         Caption         =   "Caption"
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Current Tasks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.PictureBox picF 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   2520
      ScaleHeight     =   1815
      ScaleWidth      =   1095
      TabIndex        =   9
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picTrashCan 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   240
      ScaleHeight     =   735
      ScaleWidth      =   1575
      TabIndex        =   6
      Top             =   7080
      Width           =   1575
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trash Can"
         Height          =   195
         Left            =   0
         TabIndex        =   7
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.TextBox txtRename 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox picRunMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   4560
      ScaleHeight     =   2985
      ScaleWidth      =   2625
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   2655
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close Menu"
         Height          =   255
         Left            =   1320
         TabIndex        =   2
         Top             =   2760
         Width           =   1335
      End
      Begin VB.ListBox lstRunMenu 
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
         ForeColor       =   &H00404040&
         Height          =   2460
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.PictureBox picDTIcon 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   735
      Index           =   0
      Left            =   120
      ScaleHeight     =   735
      ScaleWidth      =   1335
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblIconInfo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   720
      TabIndex        =   8
      Top             =   2400
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Menu mnuProperties 
      Caption         =   "&Object"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "&Rename"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuObjParam 
         Caption         =   "Object &Parameters..."
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Object O&ptions..."
      End
   End
   Begin VB.Menu mnuDesktop 
      Caption         =   "&Desktop"
      Begin VB.Menu mnuArrange 
         Caption         =   "&Arrange Objects"
      End
      Begin VB.Menu mnuTile 
         Caption         =   "Tile"
         Begin VB.Menu mnuTHoriz 
            Caption         =   "Horizontally"
         End
         Begin VB.Menu mnuTVert 
            Caption         =   "Vertically"
         End
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh Now"
      End
      Begin VB.Menu mnuReloadDesktop 
         Caption         =   "R&eload Desktop..."
      End
      Begin VB.Menu mnuRunProg 
         Caption         =   "R&un Program..."
      End
      Begin VB.Menu mnuCreateNew 
         Caption         =   "Create &New"
         Begin VB.Menu mnuCRef 
            Caption         =   "Reference"
         End
         Begin VB.Menu mnuCFolder 
            Caption         =   "Folder"
         End
      End
      Begin VB.Menu mnuRunList 
         Caption         =   "Classic Menu"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuToolbar 
         Caption         =   "Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuTaskList 
         Caption         =   "&Task List..."
      End
      Begin VB.Menu mnuLockDesktop 
         Caption         =   "&Lock Desktop..."
      End
      Begin VB.Menu mnuHideDesktop 
         Caption         =   "&Hide Desktop"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuControlRoom 
         Caption         =   "&Control Room..."
      End
      Begin VB.Menu mnuProfEdit 
         Caption         =   "&Profile Editor..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExitXOS 
         Caption         =   "E&xit SmartDesk..."
      End
      Begin VB.Menu mnuShutDown 
         Caption         =   "&Shut Down..."
      End
      Begin VB.Menu mnu 
         Caption         =   "&Restart..."
      End
   End
End
Attribute VB_Name = "frmLockDesktop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdRun_Click()
    
End Sub

Private Sub cmdViewTime_Click()
    frmTimelog.Show
End Sub

Private Sub cmdClose_Click()
    picRunMenu.Visible = False
   
End Sub

Private Sub cmdTask_Click(Index As Integer)
    On Error Resume Next
    NewTask2$ = cmdTask(Index).Tag
    AppActivate NewTask2$
    SendKeys "% R"
End Sub

Private Sub Command1_Click()
    picRunMenu.Visible = False
End Sub

Private Sub Form_Click()
    If szXOSVar("MISC_AUTOHIDEMENU") = "TRUE" Then
        picRunMenu.Visible = False
    End If
    'frmToolbox.Show
    'Call MakeTopWindow(frmToolbox)
End Sub

Private Sub Form_DblClick()
    frmTaskList.Show
End Sub


Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
    On Error Resume Next
    Source.Left = X
    Source.Top = Y
    
    Call ProfileWrite("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(Source.Index)), "Top", Trim$(Str$(picDTIcon(Source.Index).Top)))
    Call ProfileWrite("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(Source.Index)), "Left", Trim$(Str$(picDTIcon(Source.Index).Left)))
    lblIconInfo.Visible = False
    
    InADrag = False
End Sub


Private Sub Form_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
        On Error Resume Next
        lblIconInfo.Left = 0
        lblIconInfo.Top = 0
        lblIconInfo.Caption = "Moving '" & ProfileRead("C:\XOS\DESKTOP\DESKMENU.INF", Source.Index, "Title") & "'" & vbCrLf & "Top:  " & Y & vbCrLf & "Left:  " & X
        
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If szXOSVar("POLICY_ALLOWJUMPIN") = "TRUE" Then
        If KeyAscii = Val(szXOSVar("KEYSTROKE_JUMPIN")) Then
            frmAdminLog.Show vbModal
        End If
    End If
End Sub


Private Sub Form_Load()
    'Open szXOSVar("MISC_CBMEMFILE") For Input As #18
    'X$ = Input$(LOF(18), 18)
    'Close #18
    'Clipboard.SetText X$
    'frmToolBar.Show
    'Call MakeTopWindow(frmToolBar)
    
    
    FullQuit = True
    Call MakeTopWindow(frmToolbox)
    HelpData$ = szReadFromFile(szXOSVar("DEFAULT_XMF"))
    
    lstRunMenu.AddItem "Help..."
    
    
    
    
    If szXOSVar("MISC_LOADPICMENU") = "TRUE" Then
        lstRunMenu.AddItem "Load Picture - Tiled"
        lstRunMenu.AddItem "Load Picture - Centered"
        lstRunMenu.AddItem "Load Picture - Stretched"
        
    End If
    lstRunMenu.AddItem "View Image..."
    
    If szXOSVar("MISC_DEBUG") = "TRUE" Then
        lstRunMenu.AddItem "AbsExit"
        lstRunMenu.AddItem "Minimize"
        lstRunMenu.AddItem "Maximize"
    End If
    
    If szXOSVar("INIT_SHOWSYSMSG") = "TRUE" Then
        Open szXOSVar("MISC_SYSMSG") For Input As #11
        X$ = Input$(LOF(11), 11)
        MsgBox X$, 64, szXOSVar("MISC_SYSMSGCAPTION")
        Close #11
    End If
      
    Call GetProfileHeaders("C:\XOS\USER\CONFIG\APPS.TXT")
    
    For i = 1 To NumHeaders
        lstRunMenu.AddItem HeaderData(i)
    Next
    lstRunMenu.AddItem "View timelog..."
    If szXOSVar("POLICY_ALLOWVARCHANGE") = "TRUE" Then
    '    lstRunMenu.AddItem "Change variable..."
    End If
    
    If szXOSVar("POLICY_ALLOWSEEVAR") = "TRUE" Then
        lstRunMenu.AddItem "See variables..."
    End If
    
    If szXOSVar("POLICY_ALLOWSYSTEMDESKTOP") = "TRUE" Then
        lstRunMenu.AddItem "System desktop"
    End If
    
    If szXOSVar("POLICY_ALLOWTRUEEXIT") = "TRUE" Then
      lstRunMenu.AddItem "Exit XOS"
    End If
    
    IcnTotal = Val(ProfileRead("C:\XOS\DESKTOP\DESKMENU.INF", "Desktop", "ICTotal"))
    frmMessage.ProgressBar1.Max = IcnTotal
    PDefaultFont = szXOSVar("MISC_DESKTOPFONT")
    PDefaultSize = Val(szXOSVar("MISC_DFSIZE"))
        
        
    On Error Resume Next
        For i = 1 To IcnTotal
        PFont = PDefaultFont
        PFSize = PDefaultSize
        frmMessage.ProgressBar1.Value = i
        
        frmMessage.Refresh
        'MsgBox "'" & Trim$(Str$(i)) & "'"
        If ProfileRead("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(i)), "AutoLoad") = "FALSE" Then
            GoTo NoLoad
        End If
        pic$ = ProfileRead("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(i)), "Icon")
        PTop = Val(ProfileRead("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(i)), "Top"))
        PLeft = Val(ProfileRead("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(i)), "Left"))
        PFont = ProfileRead("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(i)), "Font")
        PFSize = Val(ProfileRead("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(i)), "Size"))
        Load picDTIcon(i)
        picDTIcon(i).Visible = True
        picDTIcon(i).Tag = ProfileRead$("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(i)), "Runs")
        If PFont = "" Then
            PFont = PDefaultFont
            PFSize = PDefaultSize
        End If
        picDTIcon(i).Font = PFont
            picDTIcon(i).FontSize = PFSize
        picDTIcon(i).Top = PTop
        picDTIcon(i).Left = PLeft
        picDTIcon(i).Picture = LoadPicture(pic$)
        picDTIcon(i).CurrentY = 500
        picDTIcon(i).Width = picDTIcon(i).TextWidth(ProfileRead$("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(i)), "Title"))
        
        picDTIcon(i).Print ProfileRead$("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(i)), "Title")
    
        frmMessage.Label1(1) = "Loaded " & picDTIcon(i).Tag
    
        PFont = PDefaultFont
        PFSize = PDefaultSize
NoLoad:
    Next
    frmMessage.Label1(0) = "SmartDesk Loaded"
    frmMessage.Label1(1) = i - 1 & " Items"
    frmMessage.Refresh
    Delay Val(szXOSVar("MISC_LOGOTIMING"))
    frmToolbox.Show
    frmToolbox.Left = 0
    frmToolbox.Top = Screen.Height - frmToolbox.Height
    Unload frmBlank
    Unload frmMessage
    frmMessage.WindowState = 1
    
    DoEvents
    Me.Refresh
    Me.SetFocus
    Me.Refresh
    'Load cmdTask(TaskIndex)
    'cmdTask(TaskIndex).Caption = Left$(ListItem$, 10)
    'cmdTask(TaskIndex).Visible = True
    'cmdTask(TaskIndex).Top = cmdTask(TaskIndex - 1).Top + cmdTask(TaskIndex - 1).Height + 100
    'cmdTask(TaskIndex).Tag = Trim$(ListItem$)
    'cmdTask(TaskIndex).ToolTipText = Trim$(ListItem$)
    TaskIndex = -1
    Dim hWnd&
    hWnd& = GetDesktopWindow()
    hWnd& = GetWindow(hWnd&, GW_CHILD)
    Do
        
        BufSize = GetWindowTextLength(hWnd&) + 1
        
        dummy = GetWindowText(hWnd&, WinText, BufSize)
        WinTextT = Trim$(WinText)
        If IsWindowVisible(hWnd&) Then
            If BufSize > 4 Then
                'MsgBox Len(Trim$(WinText))
                
                TaskIndex = TaskIndex + 1
                Load cmdTask(TaskIndex)
                cmdTask(TaskIndex).Caption = Left(Trim$(WinTextT), 20)
                cmdTask(TaskIndex).Visible = True
                cmdTask(TaskIndex).Top = cmdTask(TaskIndex - 1).Top + cmdTask(TaskIndex - 1).Height + 100
                cmdTask(TaskIndex).Tag = Trim$(WinTextT)
                cmdTask(TaskIndex).ToolTipText = Trim$(WinTextT)
                frmTaskList.lstTasks.AddItem WinTextT
            End If
        End If
        hWnd& = GetWindow(hWnd&, GW_HWNDNEXT)
        
    
    Loop While hWnd& <> 0
    
    If szXOSVar("MISC_LOADPICTURE") = "TRUE" Then
        LoadMethod = szXOSVar("MISC_PAPERMETHOD")
        frmLockDesktop.WindowState = 2
        
        frmMessage.WindowState = 1
        
        MsgBox "'" & LoadMethod & "'"
        
        
        Select Case LoadMethod
            Case "00"        'Tiled load
                picF.Picture = LoadPicture(szXOSVar("MISC_DESKTOPPICTURE"))
                For i = 0 To 2
                    For j = 0 To 2
                        frmLockDesktop.PaintPicture picF.Picture, j * picF.Width, i * picF.Height, picF.Width, picF.Height
                Next j, i
            Case "01"       'Centered load
                picF.Picture = LoadPicture(szXOSVar("MISC_DESKTOPPICTURE"))
                frmLockDesktop.PaintPicture picF.Picture, (frmLockDesktop.Width / 2) - picF.Width / 2, (frmLockDesktop.Height / 2) - picF.Height / 2
            Case "02"       'Streched load
                picF.Picture = LoadPicture(szXOSVar("MISC_DESKTOPPICTURE"))
                frmLockDesktop.PaintPicture picF.Picture, 0, 0, frmLockDesktop.ScaleWidth, frmLockDesktop.ScaleHeight
       End Select
    End If
    If szXOSVar("INIT_AUTOLOAD") = "TRUE" Then
        dummy = FreeFile
        Open "c:\xos\user\config\autoload" For Input As dummy
        Do While Not EOF(dummy)
            Line Input #dummy, Temp$
            
            X = Shell(Temp$, vbMinimizedNoFocurs)
        Loop
        Close #dummy
    End If
    'frmMeeting.Show
    
    If szXOSVar("MISC_SHOWTIP") = "TRUE" Then
        If ShowedTipOnce = False Then
            frmTip.Show vbModal
        End If
    End If
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        SaveX = X
        SaveY = Y
        PopupMenu mnuDesktop
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If InADrag = True Then
        
        
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    If FullQuit = False Then Exit Sub
    If ProfileRead("C:\XOS\USER\PROFILE", "Environment", "POLICY_REQLOG") = "TRUE" Then
        Print #1, "$"; Date$; "-END"
        Print #1, "Session completed "; Time$
        Print #1, "*END"
    End If
    Open szXOSVar("MISC_CBMEMFILE") For Output As #18
    Print #18, Clipboard.GetText
    
    frmBlank.Show
    
    
    frmMessage.Label1(0).Caption = "Closing Owned Windows"
    frmMessage.Label1(1).Caption = "Please Wait"
    frmMessage.ProgressBar1.Max = (RunIndex + 1) * 2
    frmMessage.Show
    frmMessage.Top = 0
    frmMessage.Left = 0
    frmMessage.Refresh
    
     Unload frmToolbox
    Unload frmLockDesktop
    For i = 1 To RunIndex + 1
    
        frmMessage.ProgressBar1.Value = frmMessage.ProgressBar1.Value + 1
        If RunList(i) <> 0 Then
           
            AppActivate RunList(i)
            
            SendKeys "% n", True
        End If
    Next
    For i = 1 To RunIndex + 1
    
    
        If RunList(i) <> 0 Then
           frmMessage.ProgressBar1.Value = frmMessage.ProgressBar1.Value + 1
            AppActivate RunList(i)
            
            SendKeys "%{F4}", True
        End If
    Next
    Unload frmMessage
    End
End Sub


Private Sub Form_Resize()
    picTrashCan.Top = Screen.Height - frmToolbox.Height - picTrashCan.Height - picTrashCan.TextHeight("Hello") - 100
    
    
End Sub

Private Sub imgToolbar_DblClick()
    frmToolbox.Show
    imgToolbar.Visible = False
End Sub


Private Sub lblIconInfo_Click()
    lblIconInfo.Visible = False
End Sub

Private Sub lstRunMenu_DblClick()
On Error Resume Next
    SSPanel2.BackColor = QBColor(7)
    SSPanel2.ForeColor = QBColor(0)
    If szXOSVar("MISC_AUTOHIDEMENU") = "TRUE" Then
        picRunMenu.Visible = False
    End If
    If lstRunMenu.Text = "System desktop" Then
        frmLockDesktop.WindowState = 1
    ElseIf lstRunMenu.Text = "View Image..." Then
        PicV$ = InputBox("Enter Image File:", "Image Viewer")
        frmImgViewer.picViewer.Picture = LoadPicture(PicV$)
        frmImgViewer.Show
        
        frmImgViewer.Caption = PicV$
        
        
    
    ElseIf lstRunMenu.Text = "AbsExit" Then
        Unload frmLockDesktop
        End
    ElseIf lstRunMenu.Text = "Help..." Then
        frmHelpQuery.Show
    ElseIf lstRunMenu.Text = "Load Picture - Tiled" Then
        picF.Picture = LoadPicture(szXOSVar("MISC_DESKTOPPICTURE"))
        For i = 0 To 2
        For j = 0 To 2
        frmLockDesktop.PaintPicture picF.Picture, j * picF.Width, i * picF.Height, picF.Width, picF.Height
        Next j, i
    ElseIf lstRunMenu.Text = "Load Picture - Centered" Then
        picF.Picture = LoadPicture(szXOSVar("MISC_DESKTOPPICTURE"))
        frmLockDesktop.PaintPicture picF.Picture, (frmLockDesktop.Width / 2) - picF.Width / 2, (frmLockDesktop.Height / 2) - picF.Height / 2
    ElseIf lstRunMenu.Text = "Load Picture - Stretched" Then
        picF.Picture = LoadPicture(szXOSVar("MISC_DESKTOPPICTURE"))
        frmLockDesktop.PaintPicture picF.Picture, 0, 0, frmLockDesktop.ScaleWidth, frmLockDesktop.ScaleHeight
        
    ElseIf lstRunMenu.Text = "Minimize" Then
        frmLockDesktop.WindowState = 1
    ElseIf lstRunMenu.Text = "Maximize" Then
        frmLockDesktop.WindowState = 1
    ElseIf lstRunMenu.Text = "View timelog..." Then
        frmTimelog.Show
    ElseIf lstRunMenu.Text = "Exit XOS" Then
        frmLockDesktop.Hide
    
    ElseIf lstRunMenu.Text = "See variables..." Then
        frmShowVar.Show

    Else
        EXEName$ = ProfileRead("C:\XOS\USER\CONFIG\APPS.TXT", lstRunMenu.Text, "EXEName")
        
        SMMD = Shell(EXEName$, vbNormalFocus)
        If Error$(Err) = "File not found" Then
            If ProfileRead("C:\XOS\USER\PROFILE", "Environment", "POLICY_ALLOWMENUCHANGES") = "TRUE" Then
                Ret% = MsgBox("File not found  '" & EXEName$ & "'" & vbCrLf & vbCrLf & "Would you like to change the program name?", 16 + vbYesNo, lstRunMenu.Text)
                If Ret% = vbYes Then
                    NewName$ = InputBox("Enter file name for " & lstRunMenu.Text & ":", "Update Reference", EXEName$)
                    Call ProfileWrite("C:\XOS\USER\CONFIG\APPS.TXT", lstRunMenu.Text, "EXEName", NewName$)
                End If
            Else
                MsgBox "File not found  '" & EXEName$ & "'" & vbCrLf & vbCrLf & "Contact your system administrator.", 16, lstRunMenu.Text
            End If
        End If
    End If
End Sub


Private Sub mnu_Click()
     frmMessage.Label1(0) = "System Restart"
        frmMessage.Label1(1) = "Please Wait"
        frmMessage.Show
        frmMessage.Refresh
        Delay 2
     X = ExitWindowsEx(EWX_REBOOT, 0)
     End
End Sub

Private Sub mnuArrange_Click()
    On Error Resume Next
    picDTIcon(1).Top = 100
    picDTIcon(1).Left = 100
    Call ProfileWrite("C:\XOS\DESKTOP\DESKMENU.INF", "1", "Top", Trim$(Str$(picDTIcon(1).Top)))
    Call ProfileWrite("C:\XOS\DESKTOP\DESKMENU.INF", "1", "Left", Trim$(Str$(picDTIcon(1).Left)))
    NextLeft = 100
    For i = 2 To IcnTotal
        If picDTIcon(i - 1).Top >= 5200 Then
            picDTIcon(i).Top = 100
            picDTIcon(i).Left = picDTIcon(i - 1).Left + 1435
            NextLeft = picDTIcon(i).Left
        Else
        picDTIcon(i).Top = picDTIcon(i - 1).Top + picDTIcon(i - 1).Height + 100
        
        picDTIcon(i).Left = NextLeft
        End If
        Call ProfileWrite("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(i)), "Top", Trim$(Str$(picDTIcon(i).Top)))
        Call ProfileWrite("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(i)), "Left", Trim$(Str$(picDTIcon(i).Left)))
    Next
        
End Sub

Private Sub mnuCFolder_Click()
    IcnTotal = IcnTotal + 1
    ICTotal = IcnTotal
    Call ProfileWrite("C:\XOS\DESKTOP\DESKMENU.INF", "Desktop", "ICTotal", Trim$(Str$(ICTotal)))
    Call ProfileWrite("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(ICTotal)), "LoadAuto", "TRUE")
    Call ProfileWrite("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(ICTotal)), "Left", Trim$(Str$(SaveX)))
    Call ProfileWrite("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(ICTotal)), "Top", Trim$(Str$(SaveY)))
    Call ProfileWrite("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(ICTotal)), "Title", "New Folder")
    Call ProfileWrite("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(ICTotal)), "Icon", "C:\XOS\DESKTOP\ICONS\CLSDFOLD.ICO")
    IcnTotal = ICTotal
     On Error Resume Next
        For i = 1 To IcnTotal
        'frmMessage.ProgressBar1.Value = i
        'frmMessage.Refresh
        pic$ = ProfileRead("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(i)), "Icon")
        PTop = Val(ProfileRead("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(i)), "Top"))
        PLeft = Val(ProfileRead("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(i)), "Left"))
    
        Load picDTIcon(i)
        picDTIcon(i).Visible = True
        picDTIcon(i).Tag = ProfileRead$("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(i)), "Runs")
        picDTIcon(i).Top = PTop
        picDTIcon(i).Left = PLeft
        picDTIcon(i).Picture = LoadPicture(pic$)
        picDTIcon(i).CurrentY = 500
        picDTIcon(i).Width = picDTIcon(i).TextWidth(ProfileRead$("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(i)), "Title"))
        picDTIcon(i).Print ProfileRead$("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(i)), "Title")
    Next
End Sub

Private Sub mnuControlRoom_Click()
    frmControlRoom.Show
End Sub

Private Sub mnuCRef_Click()
    IcnTotal = IcnTotal + 1
    ICTotal = IcnTotal
    Call ProfileWrite("C:\XOS\DESKTOP\DESKMENU.INF", "Desktop", "ICTotal", Trim$(Str$(ICTotal)))
    Call ProfileWrite("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(ICTotal)), "LoadAuto", "TRUE")
    Call ProfileWrite("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(ICTotal)), "Left", Trim$(Str$(SaveX)))
    Call ProfileWrite("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(ICTotal)), "Top", Trim$(Str$(SaveY)))
    Call ProfileWrite("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(ICTotal)), "Title", "New Reference")
    Call ProfileWrite("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(ICTotal)), "Icon", "C:\XOS\DESKTOP\ICONS\BOOK04.ICO")
    IcnTotal = ICTotal
     On Error Resume Next
        For i = 1 To IcnTotal
        'frmMessage.ProgressBar1.Value = i
        'frmMessage.Refresh
        pic$ = ProfileRead("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(i)), "Icon")
        PTop = Val(ProfileRead("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(i)), "Top"))
        PLeft = Val(ProfileRead("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(i)), "Left"))
    
        Load picDTIcon(i)
        picDTIcon(i).Visible = True
        picDTIcon(i).Tag = ProfileRead$("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(i)), "Runs")
        picDTIcon(i).Top = PTop
        picDTIcon(i).Left = PLeft
        picDTIcon(i).Picture = LoadPicture(pic$)
        picDTIcon(i).CurrentY = 500
        picDTIcon(i).Width = picDTIcon(i).TextWidth(ProfileRead$("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(i)), "Title"))
        picDTIcon(i).Print ProfileRead$("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(i)), "Title")
    Next
    CurrentIDX = IcnTotal
    frmRefAsst.Show
End Sub

Private Sub mnuDelete_Click()
    Call ProfileWrite("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(CurrentIDX)), "AutoLoad", "FALSE")
    Call ReloadDesktop
End Sub

Private Sub mnuExitXOS_Click()
    FullQuit = True
    Unload frmLockDesktop
End Sub

Private Sub mnuHideDesktop_Click()
    frmLockDesktop.Hide
End Sub

Private Sub mnuLockDesktop_Click()
    frmToolbox.Hide
    frmBlank.Show
    frmBlank.cmdReturn.Visible = True
End Sub

Private Sub mnuObjParam_Click()
    lblIconInfo.Left = picDTIcon(CurrentIDX).Left + picDTIcon(CurrentIDX).Width + 100
    lblIconInfo.Top = picDTIcon(CurrentIDX).Top
    lblIconInfo.Caption = "Reference:  " & picDTIcon(CurrentIDX).Tag & vbCrLf & "Item Index:  " & CurrentIDX & vbCrLf & "Top:  " & picDTIcon(CurrentIDX).Top & vbCrLf & "Left:  " & picDTIcon(CurrentIDX).Left
    
    
    lblIconInfo.Visible = True
    
End Sub

Private Sub mnuOpen_Click()
    Call picDTIcon_DblClick(CurrentIDX)
End Sub

Private Sub mnuOptions_Click()
    frmObjectOptions.Show
    frmObjectOptions.txtTitle.Text = ProfileRead("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(CurrentIDX)), "Title")
    frmObjectOptions.txtCommandLine.Text = picDTIcon(CurrentIDX).Tag
    frmObjectOptions.lblIconName.Caption = LCase$(ProfileRead("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(CurrentIDX)), "Icon"))
End Sub

Private Sub mnuProfEdit_Click()
    frmShowVar.Show
End Sub

Private Sub mnuRefresh_Click()
    Refresh
End Sub

Private Sub mnuReloadDesktop_Click()
    Call ReloadDesktop
    
End Sub

Private Sub mnuRename_Click()
    txtRename.Top = picDTIcon(CurrentIDX).Top + 500
    txtRename.Left = picDTIcon(CurrentIDX).Left
    txtRename.Text = ProfileRead$("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(CurrentIDX)), "Title")
    txtRename.Visible = True
    txtRename.SetFocus
    txtRename.SelLength = Len(txtRename.Text)
End Sub

Private Sub mnuRunList_Click()
    picRunMenu.Visible = True
    picRunMenu.Left = SaveX
    picRunMenu.Top = SaveY
End Sub

Private Sub mnuShutDown_Click()
        frmMessage.Label1(0) = "System Shutdown"
        frmMessage.Label1(1) = "Please Wait"
        frmMessage.Show
        frmMessage.Refresh
        Delay 2
        X = ExitWindowsEx(EWX_SHUTDOWN, 0)
        End
End Sub

Private Sub mnuTaskList_Click()
    frmTaskList.Show
End Sub

Private Sub mnuToolbar_Click()
    Select Case mnuToolbar.Checked
        Case True
            Unload frmToolbox
            mnuToolbar.Checked = False
        Case False
            frmToolbox.Show
            mnuToolbar.Checked = True
    End Select
End Sub

Private Sub picDTIcon_DblClick(Index As Integer)
    On Error Resume Next
    Select Case picDTIcon(Index).Tag
        Case "system.main"
               
        Case Else
            If Not Left$(picDTIcon(Index).Tag, 4) = "int:" Then
                Dim WinTextF As String * 100
                Dim TaskHWnd As Long
                Dim boom As Long
                X = Shell(picDTIcon(Index).Tag, vbNormalFocus)
                RunIndex = RunIndex + 1
                RunList(RunIndex) = X
                AppActivate X
                TaskHWnd = GetForegroundWindow()
                boom = GetWindowText(TaskHWnd, WinTextF, 100)
                RunListText(RunIndex) = Trim$(WinTextF)
            Else
                Select Case Trim$(Mid$(picDTIcon(Index).Tag, 5))
                    Case "stopwatch"
                        frmStopwatch.Show
                    Case "web"
                        dlgGetURL.Show
                    Case "drives"
                        frmSelectDrive.Show
                End Select
            End If
    End Select
End Sub

Private Sub picDTIcon_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SaveX = picDTIcon(Index).Left
    SaveY = picDTIcon(Index).Top
    If Button = 2 Then
        CurrentIDX = Index
        PopupMenu mnuProperties
        Exit Sub
    End If
    lblIconInfo.Visible = True
    picDTIcon(Index).Drag 1
    InADrag = True
    
End Sub



Private Sub picDTIcon_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'picDTIcon(Index).Drag 2
    
End Sub


Private Sub SSPanel2_Click()
    picRunMenu.Height = 0
    picRunMenu.Visible = True
    Do
        picRunMenu.Height = picRunMenu.Height + 30
        'lstRunMenu.Refresh
    Loop Until picRunMenu.Height = 2055
    SSPanel2.BackColor = QBColor(1)
    SSPanel2.ForeColor = QBColor(15)
End Sub

Private Sub tmrClock_Timer()
    pnlClock = Format$(Now, szXOSVar("LOCALE_TIMEFORMAT"))
    
End Sub


Private Sub picRunMenu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picRunMenu.Drag 1
    'lblIconInfo.Visible = True
End Sub


Private Sub picTrashCan_DblClick()
    frmTrashCan.Show
    
End Sub


Private Sub tmrTaskUpdate_Timer()
    On Error Resume Next
    For i = 1 To cmdTask.UBound
        Unload cmdTask(i)
    Next
    frmTaskList.lstTasks.Clear
    
    TaskIndex = -1
    Dim hWnd&
    hWnd& = GetDesktopWindow()
    hWnd& = GetWindow(hWnd&, GW_CHILD)
    Do
        
        BufSize = GetWindowTextLength(hWnd&) + 1
        
        dummy = GetWindowText(hWnd&, WinText, BufSize)
        WinTextT = Trim$(WinText)
        If IsWindowVisible(hWnd&) = 1 Then
            
            
            If BufSize > 4 Then
                'MsgBox Len(Trim$(WinText))
                
                TaskIndex = TaskIndex + 1
                Load cmdTask(TaskIndex)
                cmdTask(TaskIndex).Caption = Left(Trim$(WinTextT), 20)
                cmdTask(TaskIndex).Visible = True
                cmdTask(TaskIndex).Top = cmdTask(TaskIndex - 1).Top + cmdTask(TaskIndex - 1).Height + 100
                cmdTask(TaskIndex).Tag = Trim$(WinTextT)
                cmdTask(TaskIndex).ToolTipText = Trim$(WinTextT)
                frmTaskList.lstTasks.AddItem WinTextT
                
            End If
        End If
        hWnd& = GetWindow(hWnd&, GW_HWNDNEXT)
        
    
    Loop While hWnd& <> 0
End Sub

Private Sub txtRename_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtRename.Visible = False
        Call ProfileWrite("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(CurrentIDX)), "Title", txtRename.Text)
        picDTIcon(CurrentIDX).CurrentY = 500
        picDTIcon(CurrentIDX).Width = picDTIcon(i).TextWidth(ProfileRead$("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(CurrentIDX)), "Title"))
        picDTIcon(CurrentIDX).Print ProfileRead$("C:\XOS\DESKTOP\DESKMENU.INF", Trim$(Str$(CurrentIDX)), "Title")
    End If
End Sub


