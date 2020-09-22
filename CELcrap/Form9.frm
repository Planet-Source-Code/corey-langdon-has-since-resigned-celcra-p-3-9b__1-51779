VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmWebServer 
   Caption         =   "CEL-server"
   ClientHeight    =   4095
   ClientLeft      =   3615
   ClientTop       =   2115
   ClientWidth     =   8115
   LinkTopic       =   "Form9"
   ScaleHeight     =   4095
   ScaleWidth      =   8115
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Text            =   "Form9.frx":0000
      Top             =   4080
      Width           =   375
   End
   Begin VB.Frame Frame4 
      Caption         =   "Counter and Extras"
      Height          =   1335
      Left            =   2640
      TabIndex        =   15
      Top             =   2040
      Width           =   5415
      Begin VB.CheckBox Check2 
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtScript 
         Height          =   285
         Left            =   3600
         TabIndex        =   18
         Text            =   "# Visitors"
         Top             =   960
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox chkSaveCounter 
         Caption         =   "Save Counter"
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox chkCounter 
         Caption         =   "Counter on Bottom"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblScript 
         Caption         =   "Counter Script:"
         Height          =   255
         Left            =   3600
         MousePointer    =   2  'Cross
         TabIndex        =   19
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Status"
      Height          =   615
      Left            =   2640
      TabIndex        =   13
      Top             =   3480
      Width           =   5415
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Offline"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Log Settings"
      Height          =   1095
      Left            =   2640
      TabIndex        =   9
      Top             =   840
      Width           =   5415
      Begin VB.OptionButton optAddOn 
         Caption         =   "Add on"
         Enabled         =   0   'False
         Height          =   255
         Left            =   480
         TabIndex        =   12
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optOverWrite 
         Caption         =   "Overwrite"
         Enabled         =   0   'False
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox chkAutoSave 
         Caption         =   "Autosave on exit"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.Line Line10 
         X1              =   120
         X2              =   360
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line9 
         X1              =   120
         X2              =   360
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line8 
         X1              =   120
         X2              =   120
         Y1              =   480
         Y2              =   840
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Directory"
      Height          =   615
      Left            =   2640
      TabIndex        =   7
      Top             =   120
      Width           =   5415
      Begin VB.Label lblDir 
         Caption         =   "c:\InetPub"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Set Dir"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   120
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CheckBox Check1 
      Caption         =   "404's"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3840
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "0"
      Top             =   480
      Width           =   855
   End
   Begin VB.ListBox lstConnected 
      Height          =   2595
      ItemData        =   "Form9.frx":0006
      Left            =   120
      List            =   "Form9.frx":000D
      TabIndex        =   2
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   1440
      X2              =   0
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line7 
      X1              =   2520
      X2              =   2520
      Y1              =   4080
      Y2              =   0
   End
   Begin VB.Line Line6 
      X1              =   1440
      X2              =   2520
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Counter"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Line Line5 
      X1              =   1440
      X2              =   2520
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line4 
      X1              =   1440
      X2              =   0
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line3 
      X1              =   1440
      X2              =   1440
      Y1              =   840
      Y2              =   1320
   End
   Begin VB.Line Line2 
      X1              =   1440
      X2              =   1440
      Y1              =   840
      Y2              =   0
   End
End
Attribute VB_Name = "frmWebServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SelFile As String
Dim SelDat As String
Const MimeType = "text\html"

Dim Servernum As Integer
Dim data As String

Private Sub chkAutoSave_Click()
If chkAutoSave.Value = 1 Then
    optOverWrite.Enabled = True
    optAddOn.Enabled = True
Else
    optOverWrite.Enabled = False
    optAddOn.Enabled = False
End If
End Sub

Private Sub chkCounter_Click()
If chkCounter.Value = 1 Then
    lblScript.Visible = True
    txtScript.Visible = True
Else
    lblScript.Visible = False
    txtScript.Visible = False
End If
End Sub

Private Sub Command1_Click()
Winsock1(0).LocalPort = "80"
Winsock1(0).Listen
lblStatus = "Online. 0 Connections. 0KB in and 0KB out"
Command1.Enabled = False
End Sub

Private Sub Command2_Click()
For i = 0 To Winsock1.UBound
Winsock1(i).Close
Next i
Command1.Enabled = True
lblStatus = "Offline."
End Sub

Private Sub Command3_Click()
frmHTTPSet.Show
End Sub

Private Sub Label2_Click()
MsgBox "Counter Script:" & vbNewLine & "Place '#' for the count and then other text afterwards."
End Sub

Private Sub Form_Terminate()
Dim CurrentLog As String
Dim CurrentCount As String
Select Case chkAutoSave.Value
    Case 1
        If optOverWrite.Value = True Then
            Open App.Path & "\HTTPLog.txt" For Output As #1
            Print #1, Text2.Text
            Close #1
        ElseIf optAddOn.Value = True Then
            Open App.Path & "\HTTPLog.txt" For Binary As #1
            CurrentLog = String(FileLen(App.Path & "\HTTPLog.txt"), " ")
            Get #1, , CurrentLog
            Put #1, , CurrentLog & vbNewLine & vbNewLine & String(20, "-") & vbNewLine & Text2.Text & vbNewLine
            Close #1
        End If
End Select
If chkSaveCounter.Value = 1 Then
    Open App.Path & "\Counter.log" For Binary As #1
    CurrentCount = String(FileLen(App.Path & "\Counter.log"), " ")
    Get #1, , CurrentCount
    Str(CurrentCount) = Str(CurrentCount) + Str(Text1.Text)
    Put #1, , CurrentCount
    Close #1
End If
    
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Winsock1.UBound
Load Winsock1(Winsock1.UBound + 1)
Text1.Text = Text1.Text + 1
If Winsock1(Winsock1.UBound).State = sckClosed Then Winsock1(Winsock1.UBound).Close
              
Winsock1(Winsock1.UBound).Accept requestID
List1.AddItem Winsock1(Winsock1.UBound).RemoteHostIP

                                     
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Winsock1(Index).GetData HTTPdat, vbString, bytesTotal
Dim HTTPrequest As Variant
Dim FirstLine As Variant
HTTPrequest = Split(HTTPdat, vbNewLine)
FirstLine = Split(HTTPrequest(0), " ")
SelFile = Right(FirstLine(1), Len(FirstLine(1)) - 1)
Open frmHTTPSet.Dir1.Path & SelFile For Binary As #1
SelDat = String(FileLen(frmHTTPSet.Dir1.Path & "\" & SelFile), " ")
Get #1, , SelDat
Close #1
If chkCounter.Value = 0 Then
    Winsock1(Index).SendData "HTTP/1.0 200 OK" & vbCrLf & _
        "Content-Length: " & FileLen(SelDir & "\" & SelFile) & vbCrLf & _
        "Content-Type:" & MimeType & vbCrLf & vbCrLf & SelDat
ElseIf chkCounter.Value = 1 Then
        Winsock1(Index).SendData "HTTP/1.0 200 OK" & vbCrLf & _
        "Content-Length: " & FileLen(SelDir & "\" & SelFile) & vbCrLf & _
        "Content-Type:" & MimeType & vbCrLf & vbCrLf & SelDat & "<p>" & Text1.Text & " Visitors"
Text1.Text = Text1.Text + 1
End If
End Sub

Private Sub Winsock1_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock1(Index).Close
                
End Sub

Private Sub Winsock1_SendComplete(Index As Integer)
Winsock1(Index).Close

End Sub
