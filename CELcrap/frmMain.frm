VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "CELcrap 3.9"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2580
   ForeColor       =   &H0000FF00&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0CCA
   ScaleHeight     =   2250
   ScaleWidth      =   2580
   Begin VB.Label Label13 
      BackColor       =   &H00000000&
      Caption         =   "Label13"
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   1320
      Width           =   15
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      Caption         =   "Port Scanner"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   "Host Scanner"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Caption         =   "Browser"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "Host Manager"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "Web Server"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1320
      TabIndex        =   7
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "AnonyMailer"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Line Line2 
      X1              =   1320
      X2              =   1440
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      X1              =   1200
      X2              =   1200
      Y1              =   2280
      Y2              =   0
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   "DNS to IP"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "NotePad"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Chat Room"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "Port Blocker"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Web editor"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "IP Stealer"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
'DECOMMENT THE LINE BELOW FOR DISTRIBUTION!
'PLEASE GIVE CREDIT TO WWW.CEL-ERATION.COM
'YOU MAY DISTRIBUTE IF THE SETTING BELOW IS DECOMMENTED:

'Call TrialTime(Form1, "Your Trial is over.  Please purchase at www.cel-eration.com", "Trial over", vbExclamation, 15, True)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HFF00&
Label2.ForeColor = &HFF00&
Label3.ForeColor = &HFF00&
Label4.ForeColor = &HFF00&
Label5.ForeColor = &HFF00&
Label6.ForeColor = &HFF00&
Label7.ForeColor = &HFF00&
Label8.ForeColor = &HFF00&
Label9.ForeColor = &HFF00&
Label10.ForeColor = &HFF00&
Label11.ForeColor = &HFF00&
Label12.ForeColor = &HFF00&
End Sub

Private Sub Label1_Click()
frmIPStealer.Show
End Sub

Private Sub mnugggg_Click()

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HFFFF00
End Sub

Private Sub Label10_Click()
frmBrowser.Show
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label10.ForeColor = &HFFFF00
End Sub

Private Sub Label11_Click()
frmHostScan.Show
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.ForeColor = &HFFFF00
End Sub

Private Sub Label12_Click()
frmPortScan.Show
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label12.ForeColor = &HFFFF00
End Sub

Private Sub Label2_Click()
frmNotepad.Show
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &HFFFF00
End Sub

Private Sub Label3_Click()
frmWebEdit.Show
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &HFFFF00
End Sub

Private Sub Label4_Click()
Form3.Show
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &HFFFF00
End Sub

Private Sub Label5_Click()
frmChatRoom.Show
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = &HFFFF00
End Sub

Private Sub Label6_Click()
frmResolve.Show
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label6.ForeColor = &HFFFF00
End Sub

Private Sub Label7_Click()
frmAnonyMailer.Show
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label7.ForeColor = &HFFFF00
End Sub

Private Sub Label8_Click()
frmWebServer.Show
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.ForeColor = &HFFFF00
End Sub

Private Sub Label9_Click()
frmHostManage.Show
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.ForeColor = &HFFFF00
End Sub
