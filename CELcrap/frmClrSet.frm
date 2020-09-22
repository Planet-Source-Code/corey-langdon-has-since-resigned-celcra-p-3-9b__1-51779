VERSION 5.00
Begin VB.Form frmClrSet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Chat Color Settings"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3135
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picClr 
      Height          =   255
      Left            =   720
      ScaleHeight     =   195
      ScaleWidth      =   1635
      TabIndex        =   4
      Top             =   960
      Width           =   1695
   End
   Begin VB.PictureBox picPalette 
      Height          =   975
      Left            =   720
      Picture         =   "frmClrSet.frx":0000
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   109
      TabIndex        =   2
      Top             =   0
      Width           =   1695
      Begin VB.PictureBox Picture2 
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   1695
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Background"
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Font"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   1575
   End
End
Attribute VB_Name = "frmClrSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim SelColor As Boolean

Private Sub Command1_Click()
With frmChatRoom
    .Text1.ForeColor = picClr.BackColor
    .Text2.ForeColor = picClr.BackColor
    .Text3.ForeColor = picClr.BackColor
    .Text4.ForeColor = picClr.BackColor
    .Label1.ForeColor = picClr.BackColor
    .List1.ForeColor = picClr.BackColor
End With
End Sub

Private Sub Command2_Click()
With frmChatRoom
    .Text1.BackColor = picClr.BackColor
    .Text2.BackColor = picClr.BackColor
    .Text3.BackColor = picClr.BackColor
    .Text4.BackColor = picClr.BackColor
    .Label1.BackColor = picClr.BackColor
    .List1.BackColor = picClr.BackColor
    .BackColor = picClr.BackColor
End With
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub picPalette_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SelColor = True
End Sub

Private Sub picPalette_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If SelColor = True Then picClr.BackColor = GetPixel(picPalette.hdc, X, Y)
End Sub

Private Sub picPalette_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
SelColor = False
End Sub
