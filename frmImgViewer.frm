VERSION 5.00
Begin VB.Form frmImgViewer 
   Caption         =   "Document"
   ClientHeight    =   5745
   ClientLeft      =   4590
   ClientTop       =   1545
   ClientWidth     =   7380
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5745
   ScaleWidth      =   7380
   Begin VB.PictureBox picViewer 
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmImgViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    picViewer.Height = Me.ScaleHeight
    picViewer.Width = Me.ScaleWidth
End Sub


