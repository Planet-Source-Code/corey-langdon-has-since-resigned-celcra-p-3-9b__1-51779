VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmResolve 
   Caption         =   "DNS to IP"
   ClientHeight    =   1215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4110
   LinkTopic       =   "Form7"
   ScaleHeight     =   1215
   ScaleWidth      =   4110
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2520
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "255.255.255.255"
      Top             =   840
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "www.blah.com"
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmResolve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Winsock1.Connect Text1.Text, 80
End Sub


Private Sub Winsock1_Connect()
Text2.Text = Winsock1.RemoteHostIP
Winsock1.Close
End Sub

