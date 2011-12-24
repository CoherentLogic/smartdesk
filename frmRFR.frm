VERSION 5.00
Begin VB.Form frmRFR 
   Caption         =   "Rich File Reader"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4410
   LinkTopic       =   "Form2"
   ScaleHeight     =   4815
   ScaleWidth      =   4410
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picParent 
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3195
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.PictureBox picContent 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   0
         ScaleHeight     =   1215
         ScaleWidth      =   1215
         TabIndex        =   1
         Top             =   0
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmRFR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    picParent.Height = frmRFR.ScaleHeight
    picParent.Width = frmRFR.ScaleWidth
    picContent.Width = picParent.ScaleWidth
    picContent.Height = 32000
    
End Sub
