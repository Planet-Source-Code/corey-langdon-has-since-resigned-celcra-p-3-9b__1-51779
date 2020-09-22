VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmIPStealer 
   Caption         =   "IP thing"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3015
   LinkTopic       =   "Form2"
   ScaleHeight     =   6135
   ScaleWidth      =   3015
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   2985
      ItemData        =   "frmCelcrapIP.frx":0000
      Left            =   120
      List            =   "frmCelcrapIP.frx":0002
      TabIndex        =   2
      Top             =   1920
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Deactivate"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Activate"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "No Connection"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Off"
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000000&
      X1              =   1680
      X2              =   1680
      Y1              =   1920
      Y2              =   840
   End
   Begin VB.Label Label2 
      Caption         =   "Port:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000000&
      X1              =   0
      X2              =   3000
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   $"frmCelcrapIP.frx":0004
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   4920
      Width           =   2775
   End
End
Attribute VB_Name = "frmIPStealer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Winsock1.State = sckListening Then GoTo Finish
Winsock1.LocalPort = Text1.Text
Winsock1.Listen
Finish:
Label3.Caption = "On"
Label4.Caption = "No Connection"
End Sub

Private Sub Command2_Click()
Winsock1.Close
Label3.Caption = "Off"
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)

If Winsock1.State <> sckClosed Then
    Winsock1.Close
    DoEvents
    Winsock1.Accept requestID
    Label4.Caption = "Connected"
End If
If Check1.Value = 1 Then
    
    List1.AddItem Winsock1.RemoteHostIP + " Connected to you"
   
    Label4.Caption = "HTTP"
Else
    
    Label4.Caption = "No Connection"
    List1.AddItem Winsock1.RemoteHostIP + " Connected to you"
    List1.AddItem " "
    Winsock1.Close
End If

End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
If Check1.Value = 1 Then
    Winsock1.SendData "<html><center><font size=8>HTTP 404<p><br><h1><font size=2>Apache/1.32 Internal server error</font></html>"
End If
End Sub

Private Sub Winsock1_SendComplete()
 List1.AddItem "And HTTP Emulated"
 Label4.Caption = "No Connection"
 
    List1.AddItem " "
 Winsock1.Close
End Sub
