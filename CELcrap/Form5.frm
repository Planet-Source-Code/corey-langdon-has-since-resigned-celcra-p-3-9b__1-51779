VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmWebEdit 
   Caption         =   "HTML Editor - WYSIWIG"
   ClientHeight    =   8250
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10575
   LinkTopic       =   "Form5"
   ScaleHeight     =   8250
   ScaleWidth      =   10575
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10335
      ExtentX         =   18230
      ExtentY         =   6588
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4320
      Top             =   -240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   4215
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form5.frx":0000
      Top             =   4080
      Width           =   10575
   End
   Begin VB.Menu mnuSave 
      Caption         =   "Save"
   End
   Begin VB.Menu mnuCode 
      Caption         =   "Code"
      Begin VB.Menu mnuMusic 
         Caption         =   "Music"
      End
      Begin VB.Menu mnuPicture 
         Caption         =   "Picture"
      End
      Begin VB.Menu mnuLink 
         Caption         =   "Link"
      End
      Begin VB.Menu mnuBackgroundImage 
         Caption         =   "Background Image"
      End
      Begin VB.Menu mnuText 
         Caption         =   "Text"
         Begin VB.Menu mnuBold 
            Caption         =   "Bold"
         End
         Begin VB.Menu mnuItalic 
            Caption         =   "Italic"
         End
         Begin VB.Menu mnuUnderline 
            Caption         =   "Underline"
         End
         Begin VB.Menu mnuMarquee 
            Caption         =   "Marquee"
         End
      End
   End
End
Attribute VB_Name = "frmWebEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
WebBrowser1.Navigate (App.Path & "\Tmp.html")
End Sub

Private Sub Form_Load()

WebBrowser1.Navigate (App.Path & "\Tmp.html")
End Sub

Private Sub mnuBackgroundImage_Click()
Text1.Text = Text1.Text + "<body background='URL OF IMAGE'>"
End Sub

Private Sub mnuBold_Click()
Text1.Text = Text1.Text + "<b>TEXT HERE</b>"
End Sub

Private Sub mnuItalic_Click()
Text1.Text = Text1.Text + "<i>TEXT HERE</i>"
End Sub

Private Sub mnuLink_Click()
Text1.Text = Text1.Text + "<a href='URL'>TEXT</a>"
End Sub

Private Sub mnuMarquee_Click()
Text1.Text = Text1.Text + "<marquee> TEXT HERE </marquee>"
End Sub

Private Sub mnuMusic_Click()
Text1.Text = Text1.Text + "<bgsound src='URL OF MUSIC'>"
End Sub

Private Sub mnuPicture_Click()
Text1.Text = Text1.Text + "<img src='URL OF IMAGE'>"
End Sub

Private Sub mnuSave_Click()
CommonDialog1.Filter = "HTML Files (*.html)|*.html|HTM Files (*.htm)|*.htm"""
CommonDialog1.ShowSave
If CommonDialog1.FileName <> "" Then
    Open CommonDialog1.FileName For Output As #1
    Print #1, Text1.Text
    Close #1
End If
End Sub

Private Sub mnuUnderline_Click()
Text1.Text = Text1.Text + "<u> TEXT HERE</u>"
End Sub

Private Sub Text1_Change()

Open App.Path & "\Tmp.html" For Output As #1
Print #1, Text1.Text
Close #1
DoEvents
WebBrowser1.Navigate (App.Path & "\Tmp.html")
End Sub
