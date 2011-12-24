VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmShowVar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "XOS Plus - Profile Editor"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6570
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Variable Editor"
      Height          =   5655
      Left            =   4680
      TabIndex        =   4
      Top             =   240
      Width           =   4815
      Begin VB.ListBox lstStdOpt 
         Height          =   3375
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   4575
      End
      Begin VB.CommandButton cmdVarHelp 
         Caption         =   "Information"
         Height          =   375
         Left            =   3000
         TabIndex        =   9
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Add New..."
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "From File..."
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtValue 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   4215
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1320
         Top             =   3120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Standard Options:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Value:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   450
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   375
      Left            =   8160
      TabIndex        =   3
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Update"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   6000
      Width           =   1335
   End
   Begin VB.ListBox lstVarList 
      Height          =   6105
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Profile table:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   870
   End
End
Attribute VB_Name = "frmShowVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdChange_Click()
    Dim InVarname As String
    Dim InValue As String
    Dim OutVarname As String
    Dim OutValue As String
    
    InValue = txtValue.Text
    InVarname = Left$(Label2.Caption, InStr(Label2.Caption, " ") - 1)
    OutVarname = InVarname
    OutValue = InValue
    'If chkConvName.Value = 1 Then
    '    If optLCase.Value = True Then
    '        OutVarname = LCase$(InVarname)
    '    ElseIf optUCase.Value = True Then
    '        OutVarname = UCase$(InVarname)
    '    ElseIf optPreserve.Value = True Then
    '        OutVarname = InVarname
    '    End If
    'End If
    
    'If chkConvVal.Value = 1 Then
    '    If optLCase.Value = True Then
    '        OutValue = LCase$(InValue)
    '    ElseIf optUCase.Value = True Then
    '        OutValue = UCase$(InValue)
    '    ElseIf optPreserve.Value = True Then
    '        OutValue = InValue
    '    End If
    'End If
    RetVal = vbOK
    If szXOSVar("MISC_CONFIRM") = "TRUE" Then
        RetVal = MsgBox("Changing " & InVarname & "..." & vbCrLf & vbCrLf & "From " & szXOSVar(InVarname) & " to " & OutValue & vbCrLf & vbCrLf & "Please confirm...", vbOKCancel + vbQuestion, "Change Variable")
    End If
    If RetVal = vbOK Then
        Call ProfileWrite("C:\XOS\USER\PROFILE", "Environment", OutVarname, OutValue)
        If Left$(OutVarname, 6) = "MSDOS:" Then
            MsgBox "Use the XOS command 'dxlat " & Mid$(OutVarname, 7) & "' to update the DOS string table.", 64, "Profile Editor"
            
        End If
    End If
End Sub

Private Sub cmdVarHelp_Click()
    frmXSHelp.Show
    frmXSHelp.txtHelpText.Text = szExtractDataSection(HelpData$, Left$(Label2.Caption, InStr(Label2.Caption, " ") - 1))
End Sub

Private Sub Command1_Click()
    CommonDialog1.Filter = szXOSVar("MISC_DEFAULTFILTER")
    CommonDialog1.DialogTitle = "Select File..."
    CommonDialog1.ShowOpen
    
    txtValue.Text = CommonDialog1.FileName
End Sub

Private Sub Command2_Click()
    lstVarList.AddItem UCase$(InputBox$("Enter New Variable:", "Add New"))
End Sub


Private Sub Form_Load()
    On Error Resume Next
    Open "C:\XOS\USER\CONFIG\STDOPT" For Input As #16
    If szXOSVar("POLICY_ALLOWVARCHANGE") = "TRUE" Then
        cmdChange.Enabled = True
    End If
    X = FreeFile
    Open "C:\XOS\USER\PROFILE" For Input As #X
    
    Do While Not EOF(X)
        Line Input #X, Temp$
        If Temp$ = "[Environment]" Then
            Do While Temp$ <> ":_Init"
                Line Input #X, Temp$
                lstVarList.AddItem RTrim$(Left$(Temp$, InStr(Temp$, "=") - 1))
            Loop
        End If
    Loop
    Close #X
End Sub


Private Sub lstStdOpt_Click()
    txtValue.Text = lstStdOpt.Text
End Sub

Private Sub lstVarList_Click()
    txtValue.Text = szXOSVar(lstVarList.Text)
    Label2.Caption = lstVarList.Text & " Value:"
    
    lstStdOpt.Clear
    
    Do While Not EOF(16)
        Line Input #16, Temp$
        
        If Left$(Temp$, 1) = "$" Then
            If InStr(Mid$(Temp$, 2), lstVarList.Text) > 0 Then
                Do
                    Line Input #16, Temp$
                    If Temp$ = "*END" Then
                        Seek #16, 1
                        Exit Sub
                    End If
                    lstStdOpt.AddItem Temp$
                Loop
            End If
        End If
    Loop
    Seek #16, 1
End Sub


