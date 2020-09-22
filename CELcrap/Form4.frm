VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmChatRoom 
   BackColor       =   &H8000000A&
   Caption         =   "Chat Room"
   ClientHeight    =   7050
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9195
   LinkTopic       =   "Form4"
   ScaleHeight     =   7050
   ScaleWidth      =   9195
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   7800
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   8400
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      BackColor       =   &H8000000B&
      ForeColor       =   &H0000FF00&
      Height          =   2595
      ItemData        =   "Form4.frx":0000
      Left            =   120
      List            =   "Form4.frx":0002
      TabIndex        =   8
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Send"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H8000000B&
      ForeColor       =   &H0000FF00&
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "Form4.frx":0004
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000B&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "IP"
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000B&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "NickName"
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Connect"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop Server"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run Server"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000B&
      ForeColor       =   &H0000FF00&
      Height          =   6855
      Left            =   2040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Connected"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Line Line4 
      X1              =   1920
      X2              =   0
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line3 
      X1              =   1920
      X2              =   0
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line2 
      X1              =   1920
      X2              =   0
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   1920
      X2              =   1920
      Y1              =   7080
      Y2              =   0
   End
   Begin VB.Menu mnuColor 
      Caption         =   "Color Scheme"
   End
End
Attribute VB_Name = "frmChatRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim data
Dim Serving As Boolean
Dim Servernum

Private Sub Command1_Click()
Serving = True
Winsock1(Index).LocalPort = "1652"
Winsock1(Index).Listen
Command3.Enabled = False
Command1.Enabled = False
End Sub

Private Sub Command2_Click()
Serving = False
For i = 1 To Winsock1.UBound
Winsock1(i).Close
Next i
Command3.Enabled = True
Command1.Enabled = True
End Sub

Private Sub Command3_Click()
Port = "1652"
Winsock2.Connect Text3.Text, Port
Command3.Enabled = False
End Sub

Private Sub Command4_Click()
If Winsock2.State = sckClosed Then GoTo Finish
Winsock2.SendData Text2.Text + ": " + Text4.Text + vbNewLine
If Serving = True Then
    For i = 1 To Winsock1.UBound
    Winsock1(i).SendData Text2.Text + ": " + Text.Text & vbNewLine
    Next i
End If
Finish:
Text4.Text = ""
End Sub

Private Sub mnuColor_Click()
frmClrSet.Show
End Sub

Private Sub Winsock1_Close(Index As Integer)
Text1.Text = Text1.Text + vbNewLine + Winsock1(Index).RemoteHost + " Disconnected . . ." + vbNewLine
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)

Load Winsock1(Winsock1.UBound + 1)
If Winsock1(Winsock1.UBound).State = sckClosed Then Winsock1(Winsock1.UBound).Close
Winsock1(Winsock1.UBound).Accept requestID
List1.AddItem Winsock1(Winsock1.UBound).RemoteHostIP
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Winsock1(Index).GetData data, vbString, bytesTotal
Text1.Text = Text1.Text + data
For i = 1 To Winsock1.UBound
Winsock1(i).SendData data
Next i
End Sub

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox "Redo Settings"
End Sub

Private Sub Winsock2_Close()
Command3.Enabled = True
End Sub

Private Sub Winsock2_Connect()
Text1.Text = ""
Text1.Text = Text1.Text + "Connected to chat on " & Winsock2.RemoteHost
End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
Winsock2.GetData data, vbString, bytesTotal
Text1.Text = Text1.Text + data
End Sub

