VERSION 5.00
Begin VB.Form Form15 
   Caption         =   "Register"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3765
   LinkTopic       =   "Form15"
   ScaleHeight     =   2085
   ScaleWidth      =   3765
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "Keycode"
      Top             =   1440
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Reg-Username"
      Top             =   1080
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Register"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   3495
   End
   Begin VB.Label Label4 
      Caption         =   "Left"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Tries "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "There are"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    If Text1.Text = "JKCLstuff0102" And Text2.Text = "Ye597-f44" Then
        MsgBox "The name and code you entered was Correct!", vbInformation
        TrialTime Form1, "", "", "", 0, False
        Form1.Label13.Caption = "."
    Else
        MsgBox "The name and code you entered was InCorrect!", vbExclamation
    End If
End Sub

Private Sub Form_Load()

End Sub
