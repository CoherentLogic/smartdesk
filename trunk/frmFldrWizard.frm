VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Folder Wizard"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808000&
      Height          =   3495
      Left            =   120
      ScaleHeight     =   3435
      ScaleWidth      =   1755
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      Begin VB.Line Line2 
         BorderStyle     =   3  'Dot
         X1              =   1440
         X2              =   1800
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         X1              =   1440
         X2              =   1440
         Y1              =   840
         Y2              =   3000
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   960
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   960
         Left            =   120
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   840
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   720
         Top             =   1200
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Top             =   600
         Width           =   480
      End
   End
   Begin VB.Line Line3 
      X1              =   1920
      X2              =   6720
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00404080&
      Caption         =   "New Folder Wizard"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   4140
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
