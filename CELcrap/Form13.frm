VERSION 5.00
Begin VB.Form Form13 
   Caption         =   "Form13"
   ClientHeight    =   1050
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   1725
   LinkTopic       =   "Form13"
   ScaleHeight     =   1050
   ScaleWidth      =   1725
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Hide
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form12.Show
End Sub
