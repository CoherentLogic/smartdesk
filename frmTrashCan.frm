VERSION 5.00
Begin VB.Form frmTrashCan 
   Caption         =   "Trash Can"
   ClientHeight    =   8235
   ClientLeft      =   6030
   ClientTop       =   2535
   ClientWidth     =   6585
   LinkTopic       =   "Form2"
   ScaleHeight     =   8235
   ScaleWidth      =   6585
   Begin VB.ListBox lstTrashedItems 
      BackColor       =   &H00C0FFFF&
      Height          =   2595
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "frmTrashCan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    On Error Resume Next
    lstTrashedItems.Height = Me.ScaleHeight
    lstTrashedItems.Width = Me.ScaleWidth
    
End Sub


