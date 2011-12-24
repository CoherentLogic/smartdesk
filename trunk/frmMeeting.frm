VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{A8B3B723-0B5A-101B-B22E-00AA0037B2FC}#1.0#0"; "grid32.ocx"
Begin VB.Form frmMeeting 
   Caption         =   "SmartDesk Meeting Minder 97"
   ClientHeight    =   8235
   ClientLeft      =   4275
   ClientTop       =   2115
   ClientWidth     =   6585
   LinkTopic       =   "Form2"
   ScaleHeight     =   8235
   ScaleWidth      =   6585
   Begin Threed.SSPanel pnlToolbar 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6585
      _Version        =   65536
      _ExtentX        =   11615
      _ExtentY        =   1085
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
   End
   Begin MSGrid.Grid grdMeeting 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   3375
      _Version        =   65536
      _ExtentX        =   5953
      _ExtentY        =   6165
      _StockProps     =   77
      BackColor       =   16777215
      Rows            =   100
      Cols            =   100
   End
End
Attribute VB_Name = "frmMeeting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    On Error Resume Next
    grdMeeting.Width = frmMeeting.ScaleWidth
    grdMeeting.Height = Me.ScaleHeight - pnlToolbar.Height - 10
    
grdMeeting.Top = pnlToolbar.Height + 10

End Sub


