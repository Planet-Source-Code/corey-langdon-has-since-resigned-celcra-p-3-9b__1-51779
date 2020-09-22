VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmHostScan 
   Caption         =   "Host Scanner"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3015
   LinkTopic       =   "Form12"
   ScaleHeight     =   5295
   ScaleWidth      =   3015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Scan"
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   240
      Width           =   3015
   End
   Begin VB.ListBox List1 
      Height          =   2985
      ItemData        =   "frm12.frx":0000
      Left            =   0
      List            =   "frm12.frx":0002
      TabIndex        =   9
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Timer Delayer 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1200
      Top             =   2760
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Text            =   "Port"
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save List"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   5040
      Width           =   3015
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Random Scan"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   720
      Width           =   1575
   End
   Begin VB.OptionButton Option2 
      Caption         =   "IP range"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Text            =   "255"
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "255"
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Text            =   "255"
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Text            =   "1"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Text            =   "255"
      Top             =   1200
      Width           =   495
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4680
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Waiting"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3000
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   3000
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "frmHostScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Searching As Boolean
Dim IPnum As String
Private Sub Command1_Click()
If Searching = False Then Delayer.Interval = InputBox("How many milliseconds shall we delay connections?")
Searching = True
Dim Dig1 As Integer
Dim Dig2 As Integer
Dim Dig3 As Integer
Dim Dig4 As Integer
Dim CheckDig As Integer
If Option1.Value = True Then
    Randomize
    Value = (255 * Rnd)
    Dig1 = Str$(Value)

    Randomize
    Value = (255 * Rnd)
    Dig2 = Str$(Value)

    Randomize
    Value = (255 * Rnd)
    Dig3 = Str$(Value)

    Randomize
    Value = (255 * Rnd)
    Dig4 = Str$(Value)
    Label1.Caption = Dig1 & "." & Dig2 & "." & Dig3 & "." & Dig4
    Winsock1.Connect Label1.Caption, Text1.Text
    Delayer.Enabled = True
    DoEvents: DoEvents: DoEvents: DoEvents
ElseIf Option2.Value = True Then
    Text5.Text = Text5.Text + 1
    Dig1 = Text2.Text
    Dig2 = Text3.Text
    Dig3 = Text4.Text
    Dig4 = Text5.Text
    CheckDig = Text6.Text
    If Dig4 = CheckDig Then
        Text5.Text = "1"
        Text6.Text = "255"
        If Text4.Text < 255 Then Text4.Text = Text4.Text + 1
    End If
    If Text4.Text = "255" Then
        Text5.Text = "1"
        Text6.Text = "255"
        Text4.Text = "1"
        Text3.Text = Text3.Text + 1
    End If
    If Text3.Text = "255" Then
        Text5.Text = "1"
        Text6.Text = "255"
        Text4.Text = "1"
        Text3.Text = "1"
        Text2.Text = Text2.Text + 1
    End If
    If Text2.Text = "255" Then
        Call Command2_Click
        Label1.Caption = "Waiting"
    End If
    Label1.Caption = Dig1 & "." & Dig2 & "." & Dig3 & "." & Dig4
    Winsock1.Connect Label1.Caption, Text1.Text
    Delayer.Enabled = True
    DoEvents: DoEvents: DoEvents: DoEvents
End If
End Sub



Private Sub Command2_Click()
Winsock1.Close
Delayer.Enabled = False
Label1.Caption = "Waiting"
Option1.Value = False
Option2.Value = False
Searching = False
End Sub

Private Sub Delayer_Timer()
Winsock1.Close
Delayer.Enabled = False
Call Command1_Click
End Sub

Private Sub Form_Load()
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Searching = False
End Sub

Private Sub List1_DblClick()
List1.RemoveItem (List1.Index)
End Sub

Private Sub Option1_Click()
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
End Sub

Private Sub Option2_Click()
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
End Sub

Private Sub Winsock1_Connect()
List1.AddItem Label1.Caption & " : " & Text1.Text
Winsock1.Close
Delayer.Enabled = False
Call Command1_Click
End Sub

Private Sub Command3_Click()
Open App.Path & "\Host-" & InputBox("Create a name for this list:") & ".txt" For Output As #1
Dim ListStuff As String
For i = 0 To List1.ListCount - 1
ListStuff = ListStuff + List1.List(i) & vbNewLine
Next i
Print #1, "Hosts Found On Port " & Text1.Text & vbNewLine & Date & vbNewLine & Time & vbNewLine & ListStuff
Close #1
End Sub
