VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmBrowser 
   Caption         =   "CEL Simple Browser"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10350
   LinkTopic       =   "Form11"
   ScaleHeight     =   8670
   ScaleWidth      =   10350
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkBlockPop 
      Caption         =   "Block Popups"
      Height          =   255
      Left            =   4680
      TabIndex        =   9
      Top             =   0
      Width           =   1335
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser2 
      Height          =   135
      Left            =   10320
      TabIndex        =   8
      Top             =   4440
      Visible         =   0   'False
      Width           =   135
      ExtentX         =   238
      ExtentY         =   238
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3240
      TabIndex        =   7
      Text            =   "Keywords"
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Search"
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">>>>"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Refresh"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<<<<"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GO!"
      Height          =   255
      Left            =   9120
      TabIndex        =   2
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6000
      TabIndex        =   1
      Top             =   0
      Width           =   3135
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   10335
      ExtentX         =   18230
      ExtentY         =   14420
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   4560
      X2              =   4560
      Y1              =   360
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   2400
      X2              =   2400
      Y1              =   360
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   0
      X2              =   10320
      Y1              =   360
      Y2              =   360
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
WebBrowser1.Navigate Text1.Text
End Sub

Private Sub Command2_Click()
WebBrowser1.GoBack
End Sub

Private Sub Command3_Click()
WebBrowser1.Refresh
End Sub

Private Sub Command4_Click()
WebBrowser1.GoForward
End Sub

Private Sub Command5_Click()
WebBrowser1.Navigate ("http://www.google.com/search?hl=en&ie=UTF-8&oe=UTF-8&q=" + Text2.Text + "&btnG=Google+Search")
End Sub

Private Sub Form_Resize()
WebBrowser1.Width = Form11.Width - 100
WebBrowser1.Height = Form11.Height - 1000
Text1.Width = Form11.Width / 2 - 1800
Command1.Left = Text1.Left + Text1.Width
Line1.X2 = Form11.Width
End Sub

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
If chkBlockPop.Value = 1 Then Set ppDisp = WebBrowser2.object

End Sub

