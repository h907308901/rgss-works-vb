VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RGSS������"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4740
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   4740
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3000
      Top             =   480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   1080
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "����"
      Enabled         =   0   'False
      Height          =   3495
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   4455
      Begin VB.CommandButton Command7 
         Caption         =   "���ӽ�Ǯ"
         Height          =   375
         Left            =   1560
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "ȫ��ظ�"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "��ͼǿ�ƴ浵"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "����DEBUG"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "ǿ�ƽ���ս��"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "ִ��"
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   720
         TabIndex        =   6
         Top             =   2520
         Width           =   3615
      End
      Begin VB.Label Label5 
         Caption         =   "�ű���"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   2640
         Width           =   615
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2400
      Top             =   480
   End
   Begin VB.Label Label4 
      Caption         =   "��������"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "����ID��"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "���⣺"
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "�����"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.Image imgDrag1 
      Height          =   420
      Left            =   240
      Picture         =   "Form1.frx":5C12
      Top             =   240
      Width           =   465
   End
   Begin VB.Image imgDrag2 
      Height          =   420
      Left            =   240
      Picture         =   "Form1.frx":63D4
      Top             =   240
      Width           =   465
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim whwnd As Long, wtext As String, pid As Long, pname As String
Dim hProcess As Long, pfnRGSSEval As Long

Private Sub Command1_Click()
    Dim sPath As String, sINI As String, sDLLName As String
    On Error GoTo ErrLine
    If hProcess = 0 Then
        If pid = 0 Then
            MsgBox "����ѡ����Ϸ���ڣ�"
            Exit Sub
        End If
        sPath = GetProcessNameByProcessId(pid)
        sINI = sPath
        Mid(sINI, InStrRev(sINI, ".")) = ".ini"
        sDLLName = String(256, vbNullChar)
        GetPrivateProfileString "Game", "Library", "", sDLLName, 256, sINI
        If sDLLName = vbNullString Then GoTo BadINI
Continue1:
        hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, pid)
        pfnRGSSEval = GetRGSSEvalAddress(hProcess, sDLLName)
        If pfnRGSSEval = 0 Then
            MsgBox "��ȡRGSSEval��ַʧ�ܣ�"
            CloseHandle hProcess
            hProcess = 0
            Exit Sub
        End If
        imgDrag1.Enabled = False
        Frame1.Enabled = True
        Command1.Caption = "�ر�"
        Timer2.Enabled = True
    Else
        CloseHandle hProcess
        hProcess = 0
        imgDrag1.Enabled = True
        Frame1.Enabled = False
        Command1.Caption = "����"
        Timer2.Enabled = False
        pid = 0
        Label1 = "�����"
        Label2 = "���⣺"
        Label3 = "����ID��"
        Label4 = "��������"
    End If
    Exit Sub
ErrLine:
    MsgBox "����"
    Exit Sub
BadINI:
    If MsgBox("��ȡ��Ϸ����ʧ�ܣ��Ƿ��ֶ�������ϷRGSS������DLL����", vbYesNo) = vbYes Then
        sDLLName = InputBox("��ϷRGSS������DLL���ƣ�")
        If sDLLName <> vbNullString Then GoTo Continue1
    End If
End Sub

Private Sub Command2_Click()
    Dim script As String
    script = Text1
    CloseHandle RemoteCall(hProcess, pfnRGSSEval, script)
    SetForegroundWindow whwnd
End Sub

Private Sub Command3_Click()
    Text1 = "$game_temp.battle_abort=true"
End Sub

Private Sub Command4_Click()
    Text1 = "$DEBUG=true"
End Sub

Private Sub Command5_Click()
    Text1 = "$game_temp.save_calling=true"
End Sub

Private Sub Command6_Click()
    Text1 = "$game_system.map_interpreter.iterate_actor(0) {|actor| actor.recover_all}"
End Sub

Private Sub Command7_Click()
    Dim s As String
    s = InputBox("��Ǯ����")
    If s = vbNullString Or Not IsNumeric(s) Then Exit Sub
    Text1 = "$game_party.gain_gold(" & s & ")"
End Sub

Private Sub imgDrag1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgDrag1.Visible = False
    MousePointer = 99
    MouseIcon = LoadResPicture(101, 2)
    Timer1.Enabled = True
End Sub

Private Sub imgDrag1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgDrag1.Visible = True
    MousePointer = 0
    Timer1.Enabled = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call Command2_Click
End Sub

Private Sub Timer1_Timer()
    Dim point As POINTAPI
    GetCursorPos point
    whwnd = WindowFromPoint(point.x, point.y)
    wtext = Space(GetWindowTextLength(whwnd) + 1)
    GetWindowText whwnd, wtext, Len(wtext)
    GetWindowThreadProcessId whwnd, pid
    pname = GetFileName(GetProcessNameByProcessId(pid))
    Label1 = "�����" & whwnd
    Label2 = "���⣺" & wtext
    Label3 = "����ID��" & pid
    Label4 = "��������" & pname
End Sub

Private Sub Timer2_Timer()
    If WaitForSingleObject(hProcess, 1) = 0 Then Call Command1_Click
End Sub
