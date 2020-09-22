VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmPortScan 
   BackColor       =   &H00000000&
   Caption         =   "Port Scanner"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4035
   LinkTopic       =   "Form14"
   ScaleHeight     =   3000
   ScaleWidth      =   4035
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Clear List"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save List"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Text            =   "65355"
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Text            =   "1"
      Top             =   1680
      Width           =   735
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   1320
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Text            =   "IP"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Scan"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2985
      Left            =   1920
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label lblPort 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Current Port"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   " to"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   1680
      Width           =   255
   End
End
Attribute VB_Name = "frmPortScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Scanning As Boolean
Dim Port

Private Sub Command1_Click()
Scanning = True
Command1.Enabled = False
Command3.Enabled = False
For i = Text2.Text To Text3.Text
lblPort.Caption = i
If Scanning = False Then GoTo Finish
If Winsock1(Winsock1.UBound).State <> sckConnected Then Unload Winsock1(Winsock1.UBound)
Call PortScan(Text1.Text, lblPort.Caption)
Next i
Finish:
lblPort.Caption = "Scan Complete"
End Sub

Private Sub Command2_Click()
Scanning = False
Command1.Enabled = True
Command3.Enabled = True
End Sub

Private Sub Command3_Click()

Dim fFile As Integer
fFile = FreeFile
Open App.Path & "\Port-" & Text1.Text & ".txt" For Output As #fFile
Dim ListStr As String
ListStr = ""
For i = 0 To List1.ListCount - 1
ListStr = ListStr & Text1.Text & " was open on port " & List1.List(i) & vbNewLine
Next i
Print #fFile, ListStr
Close #fFile
End Sub

Private Sub Command4_Click()
If MsgBox("Do you want to clear this list?", vbYesNo, "Clear List?") = vbNo Then GoTo Finish
For i = 0 To List1.ListCount - 1
List1.RemoveItem (i)
Next i
Finish:
End Sub

Private Sub Form_Load()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Winsock1_Connect(Index As Integer)
List1.AddItem Winsock1(Index).RemotePort
Unload Winsock1(Index)
End Sub

Function PortScan(IP As String, Port As String)
Load Winsock1(Winsock1.UBound + 1)
Winsock1(Winsock1.UBound).Connect IP, Port
End Function
