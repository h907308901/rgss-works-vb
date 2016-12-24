VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RGSSAD Unpacker"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   5910
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command5 
      Caption         =   "猜解"
      Height          =   375
      Left            =   5160
      TabIndex        =   12
      Top             =   1560
      Width           =   495
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   2520
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Command4 
      Caption         =   "退出"
      Height          =   375
      Left            =   4680
      TabIndex        =   10
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "输出日志"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3120
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "解包"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   2
      Text            =   "DEADCAFE"
      Top             =   1560
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   960
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label4 
      Caption         =   "警告：错误的MagicKey可能会导致程序崩溃！"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   2160
      Width           =   5295
   End
   Begin VB.Label Label3 
      Caption         =   "MagicKey"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "输出路径："
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "加密档案路径："
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    Check1.Enabled = False
    ProgressBar1.Visible = True
    If ExtractFile(Text1, Text2, Val("&H" & Text3), Check1.Value <> 0, ProgressBar1) = 0 Then
        MsgBox "成功！"
    Else
        MsgBox "失败！"
    End If
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = True
    Command5.Enabled = True
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Check1.Enabled = True
    ProgressBar1.Visible = False
End Sub

Private Sub Command2_Click()
    Dim s As String
    s = ShowDialogFile(Me.hWnd, 1, "请选择加密档案", "", "加密档案文件" & Chr(0) & "*.rgssad", "", "")
    If s <> "" Then
        Text1 = s
        Text2 = Left(s, InStrRev(s, "\"))
    End If
End Sub

Private Sub Command3_Click()
    Dim s As String
    s = GetFolderPath("选择输出文件夹", Me.hWnd)
    If s <> "" Then Text2 = s
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Command5_Click()
    If MsgBox("MagicKey猜解是通过通常情况下加密档案中第一个文件为Data\Actors.rxdata工作的，存在一定不确定性，是否继续？", vbYesNo) = vbYes Then
        Text3 = Hex(MagicKeyGuess(Text1))
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = "RGSSAD Unpacker v" & CStr(App.Major) & "." & CStr(App.Minor)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKey0 To vbKey9, vbKeyA To vbKeyF, vbKeyBack
    Case Asc("a") To Asc("f")
        KeyAscii = KeyAscii - Asc("a") + vbKeyA
    Case Else
        KeyAscii = 0
    End Select
End Sub
