VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAnonyMailer 
   Caption         =   "AnonyMailer"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6855
   LinkTopic       =   "Form8"
   ScaleHeight     =   5730
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Attach"
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   4800
      TabIndex        =   9
      Top             =   1800
      Width           =   1935
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5280
      Top             =   5760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   5160
      TabIndex        =   6
      Text            =   "Masked IP"
      Top             =   960
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   4800
      TabIndex        =   5
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4800
      TabIndex        =   4
      Text            =   "SMTP server"
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send"
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6840
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text3 
      Height          =   4335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   840
      Width           =   4455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Subject"
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Your Fake email address"
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   $"Form8.frx":0000
      Height          =   2055
      Left            =   4800
      TabIndex        =   13
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Attachments:"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Connect to your ISP's SMTP server. "
      Height          =   495
      Left            =   4800
      TabIndex        =   10
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Line Line5 
      X1              =   4680
      X2              =   6840
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label3 
      Caption         =   "Status:Waiting"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   5400
      Width           =   6615
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   6840
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Label Label2 
      Caption         =   "Reciever"
      Height          =   255
      Left            =   4800
      TabIndex        =   7
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Line Line3 
      X1              =   4680
      X2              =   6840
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line2 
      X1              =   4680
      X2              =   6840
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line1 
      X1              =   4680
      X2              =   4680
      Y1              =   5280
      Y2              =   0
   End
End
Attribute VB_Name = "frmAnonyMailer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
CommonDialog1.Filter = "Any type of file (*.*)|*.*"
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Input As ATTACHMENT
    Label1.Caption = Label1.Caption + CommonDialog1.FileName
End If
End Sub

Private Sub Command2_Click()
If Winsock1.State <> sckClosed Then Winsock1.Close

Winsock1.Connect Text4.Text, "25"
DoEvents: DoEvents: DoEvents: DoEvents
Label3.Caption = "Status: Connecting"
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Winsock1_Connect()
        Label3.Caption = "Status: Sending"
        Winsock1.SendData "HELO " + Winsock1.LocalIP + vbCrLf
        Winsock1.SendData "MAIL FROM:<" + Text1.Text + " > " + vbCrLf
        Winsock1.SendData "RCPT TO:<" + Text6.Text + " > " + vbCrLf
        Winsock1.SendData "DATA" + vbCr + Lf
        Winsock1.SendData "From: <" + Text1.Text + ">" + vbCrLf + _
        "To: " + Text6.Text + vbCrLf + _
        "Subject: " + Text2.Text + vbCrLf + _
        "X-Mailer: X" + vbCrLf + _
        "Mime-Version: 1.0" + vbCrLf + _
        "Content-Type: text/html" + vbTab + "charset=us-ascii" + vbCrLf + vbCrLf & Text3.Text
        Winsock1.SendData vbCrLf + "." + vbCrLf
    Winsock1.SendData "QUIT"
    MsgBox "Mail has been sent succesfully!", vbInformation, "Mail has been sent succesfully!"
Label3.Caption = "Status:Waiting"
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Err.Clear
MsgBox "The SMTP server wont let you send through them."
End Sub
