VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmHelpQuery 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "XOS Plus - Help Query"
   ClientHeight    =   3315
   ClientLeft      =   1785
   ClientTop       =   2160
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3315
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Open"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse..."
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtHFile 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Text            =   "default.xmf"
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "&Display"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.ListBox lstContents 
      BackColor       =   &H00C0FFFF&
      Height          =   2205
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   3735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4080
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5280
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Book:"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   420
   End
End
Attribute VB_Name = "frmHelpQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ReadTopics(HFName As String)
    X = FreeFile
    Open HFName For Input As X
    
    Do While Not EOF(X)
        Line Input #X, Temp$
        If Left$(Temp$, 1) = "$" Then
            lstContents.AddItem Mid$(Temp$, 2)
        End If
    Loop
    
    Close X
End Sub


Private Sub cmdBrowse_Click()
    frmHelpBrowser.Show
End Sub

Private Sub cmdShow_Click()
    HelpData$ = szReadFromFile(szXOSVar("DEFAULT_XMF") & "\" & txtHFile.Text)
    frmXSHelp.Show
    frmXSHelp.Caption = "XOS Plus - Help on " & lstContents.Text
    frmXSHelp.txtHelpText.Text = szExtractDataSection(HelpData$, lstContents.Text)
    Unload Me
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    ReadTopics szXOSVar("DEFAULT_XMF") & "\" & txtHFile.Text
End Sub

Private Sub Form_Load()
    ReadTopics szXOSVar("DEFAULT_XMF") & "\" & txtHFile.Text
End Sub


