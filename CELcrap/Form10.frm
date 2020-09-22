VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmHostManage 
   Caption         =   "Hosts Manager"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3720
   LinkTopic       =   "Form10"
   ScaleHeight     =   4590
   ScaleWidth      =   3720
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2040
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add New Mask"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   4215
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form10.frx":0000
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "frmHostManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
URL = InputBox("What is the Website?")
IP = InputBox("What is the ip?")
Text1.Text = Text1.Text + vbNewLine + IP + "  " + URL
End Sub

Private Sub Command2_Click()
If MsgBox("Save as current 'HOSTS' file?", vbYesNo, "Save . . .?") = vbNo Then GoTo JustSave
Open "C:\WINDOWS\hosts" For Output As #1
Print #1, Text1.Text
Close #1
JustSave:
CommonDialog1.Filter = "File (*.*)|*.*"
CommonDialog1.ShowSave
If CommonDialog1.FileName <> "" Then
Open CommonDialog1.FileName For Output As #1
    Print #1, Text1.Text
    Close #1
End If
End Sub


Private Sub Form_Load()
Text1.Text = Text1.Text + vbCrLf
End Sub
