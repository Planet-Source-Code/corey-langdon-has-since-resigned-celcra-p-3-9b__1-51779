VERSION 5.00
Begin VB.Form frmHTTPSet 
   Caption         =   "CEL-Server Settings"
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3570
   LinkTopic       =   "Form2"
   ScaleHeight     =   2565
   ScaleWidth      =   3570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Update Directory"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   3435
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3375
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3600
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label1 
      Caption         =   "Directory:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmHTTPSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SelDir = Dir1.Path
MsgBox SelDir
Me.Hide
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub
