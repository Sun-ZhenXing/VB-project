VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "HTML_get"
   ClientHeight    =   5085
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   8415
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   6720
      Top             =   0
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      ExtentX         =   11668
      ExtentY         =   6800
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
      Location        =   "http:///"
   End
   Begin VB.Menu MFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu MChoice 
         Caption         =   "选项(&C)"
      End
      Begin VB.Menu MTest 
         Caption         =   "测试(&T)"
         Shortcut        =   ^T
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xxx As Integer

Private Sub Form_Load()
WebBrowser1.Navigate "about:blank"
xxx = 0
End Sub

Private Sub Form_Resize()
WebBrowser1.Width = Form1.Width - 200
WebBrowser1.Height = Form1.Height - 800
End Sub

Private Sub MTest_Click()
WebBrowser1.Document.Write "你好"
End Sub

Private Sub Timer1_Timer()
xxx = xxx + 1
If xxx = 2 Then
    WebBrowser1.Navigate Command
End If
If xxx = 6 Then
    Open "D:\html_out.txt" For Output As #1
    Print #1, WebBrowser1.Document.documentElement.outerHTML
    Close #1
    End
End If
End Sub

Private Sub WebBrowser1_DownloadBegin()
WebBrowser1.Silent = True
End Sub
