VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form Form3 
   Caption         =   "Effortless Firewall"
   ClientHeight    =   1080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7035
   LinkTopic       =   "Form3"
   ScaleHeight     =   1080
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock4 
      Left            =   1920
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock3 
      Left            =   1320
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   720
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   3360
      TabIndex        =   10
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   2280
      TabIndex        =   9
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   615
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Text            =   "Port4"
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Text            =   "Port2"
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Port3"
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Port1"
      Top             =   120
      Width           =   975
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   6960
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label3 
      Caption         =   "Socket Handles"
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Line Line1 
      X1              =   6960
      X2              =   0
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label2 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5880
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Total Stopped : "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3240
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Winsock1.LocalPort = Text1.Text
Winsock1.Listen
Winsock2.LocalPort = Text2.Text
Winsock2.Listen
Winsock3.LocalPort = Text3.Text
Winsock3.Listen
Winsock4.LocalPort = Text4.Text
Winsock4.Listen
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Bind Winsock1.LocalPort, Winsock1.LocalIP
Label2.Caption = Label2.Caption + 1
Text5.Text = Winsock1.SocketHandle
End Sub

Private Sub Winsock2_ConnectionRequest(ByVal requestID As Long)
Winsock2.Bind Winsock2.LocalPort, Winsock2.LocalIP
Label2.Caption = Label2.Caption + 1
Text6.Text = Winsock2.SocketHandle
End Sub

Private Sub Winsock3_ConnectionRequest(ByVal requestID As Long)
Winsock3.Bind Winsock3.LocalPort, Winsock3.LocalIP
Label2.Caption = Label2.Caption + 1
Text7.Text = Winsock3.SocketHandle
End Sub

Private Sub Winsock4_ConnectionRequest(ByVal requestID As Long)
Winsock4.Bind Winsock4.LocalPort, Winsock4.LocalIP
Label2.Caption = Label2.Caption + 1
Text8.Text = Winsock4.SocketHandle
End Sub

