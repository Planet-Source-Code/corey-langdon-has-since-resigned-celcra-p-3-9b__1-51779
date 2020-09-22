VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNotepad 
   Caption         =   "My NotePad"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8175
   LinkTopic       =   "Form6"
   ScaleHeight     =   5505
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6720
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "text"
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Copy All"
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Open"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   4935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   7935
   End
   Begin VB.Line Line2 
      X1              =   5520
      X2              =   5520
      Y1              =   480
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   2760
      X2              =   2760
      Y1              =   480
      Y2              =   0
   End
End
Attribute VB_Name = "frmNotepad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CommonDialog1.Filter = "Text Files (*.txt)|*.txt"
CommonDialog1.ShowSave
If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Output As FreeFile
    Print FreeFile, Text1.Text
    Close FreeFile
End If
End Sub

Private Sub Command2_Click()
CommonDialog1.Filter = "Text Files (*.txt)|*.txt"
Commonialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Input As 1
    Text1.Text = 1
    Close 1
End If
End Sub

Private Sub Command3_Click()
Text1.Text = ""
End Sub

Private Sub Dir1_Change()

End Sub

Private Sub Drive1_Change()

End Sub

Private Sub File1_Click()

End Sub

Private Sub Command4_Click()
Clipboard.SetText Text1.Text, ASCII
End Sub

