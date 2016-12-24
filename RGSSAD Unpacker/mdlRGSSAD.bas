Attribute VB_Name = "mdlRGSSAD"
Option Explicit
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (lpString As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Public Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Any, ByVal dwMoveMethod As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpWideCharStr As Any, ByVal cchWideChar As Long) As Long
Public Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpWideCharStr As Any, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, lpDefaultChar As Any, ByVal lpUsedDefaultChar As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SHCreateDirectoryEx Lib "SHELL32.DLL" Alias "SHCreateDirectoryExA" (ByVal hWnd As Long, ByVal pszPath As String, psa As Any) As Long
Public Const FILE_ALL_ACCESS = &H1F01FF
Public Const OPEN_EXISTING = 3
Public Const CREATE_ALWAYS = 2
Public Const FILE_BEGIN = 0
Public Const FILE_CURRENT = 1
Public Const FILE_END = 2
Public Const CP_UTF8 = 65001
Public Const CP_ACP = 0
Public Type RGSSAD_SUB_INFO
    FileNameSize As Long
    Offset As Long
    FileSize As Long
    MagicKey As Long
    FileName(255) As Byte
End Type
Dim bCode() As Byte

Public Function ExtractFile(FileName As String, Path As String, ByVal MagicKey As Long, ByVal Log As Boolean, Optional ByVal Progress As ProgressBar = Nothing) As Long
    Dim ret As Long
    Dim hFile As Long, hLogFile As Long, hNewFile As Long
    Dim nBytesRead As Long
    Dim Flag1 As Long, Flag2 As Long
    Dim tmpInfo As RGSSAD_SUB_INFO
    Dim InfoList() As RGSSAD_SUB_INFO, InfoCount As Long
    Dim IsEof As Boolean
    Dim i As Long, j As Long, s As String, l As Long
    Dim byt(255) As Byte, buff(1023) As Byte, leftsize As Long
    PatchCode
    'Create logging file
    If Log Then
        hLogFile = CreateFile(Path & "\rgssad_unpacker.log", FILE_ALL_ACCESS, 0, ByVal 0, CREATE_ALWAYS, 0, 0)
        If hLogFile = -1 Then 'failed
            ExtractFile = -1
            GoTo ErrLine
        End If
    End If
    WriteLog hLogFile, "===== RGSSAD Unpacker v" & CStr(App.Major) & "." & CStr(App.Minor) & " ====="
    WriteLog hLogFile, "RGSSAD File: " & FileName
    'Open RGSSAD file
    hFile = CreateFile(FileName, FILE_ALL_ACCESS, 0, ByVal 0, OPEN_EXISTING, 0, 0)
    If hFile = -1 Then 'failed
        WriteLog hLogFile, "Open file '" & FileName & "' failed"
        ExtractFile = -1
        GoTo ErrLine
    End If
    'Verify file header
    ReadFile hFile, Flag1, 4, nBytesRead, ByVal 0
    ReadFile hFile, Flag2, 4, nBytesRead, ByVal 0
    WriteLog hLogFile, "Header: " & Hex(Flag1) & " " & Hex(Flag2)
    If Flag1 <> &H53534752 Or Flag2 <> &H1004441 Then
        WriteLog hLogFile, "Not a valid RGSSAD file"
        ExtractFile = -1
        GoTo ErrLine
    End If
    'Generate file information list
    ret = ReadFile(hFile, tmpInfo.FileNameSize, 4, nBytesRead, ByVal 0)
    If ret = 1 And nBytesRead < 4 Then IsEof = True
    Do Until IsEof
        tmpInfo.FileNameSize = tmpInfo.FileNameSize Xor MagicKey
        MagicKey = MKTransform(MagicKey) 'MagicKey = MagicKey * 7 + 3
        ret = ReadFile(hFile, tmpInfo.FileName(0), tmpInfo.FileNameSize, nBytesRead, ByVal 0)
        If ret = 1 And nBytesRead < tmpInfo.FileNameSize Then IsEof = True
        For i = 0 To tmpInfo.FileNameSize - 1
            tmpInfo.FileName(i) = tmpInfo.FileName(i) Xor (MagicKey And &HFF)
            MagicKey = MKTransform(MagicKey) 'MagicKey = MagicKey * 7 + 3
        Next
        tmpInfo.FileName(i) = 0
        MultiByteToWideChar CP_UTF8, 0, tmpInfo.FileName(0), -1, byt(0), 256
        WideCharToMultiByte CP_ACP, 0, byt(0), -1, tmpInfo.FileName(0), 256, ByVal 0, 0
        ret = ReadFile(hFile, tmpInfo.FileSize, 4, nBytesRead, ByVal 0)
        If ret = 1 And nBytesRead < 4 Then IsEof = True
        tmpInfo.FileSize = tmpInfo.FileSize Xor MagicKey
        MagicKey = MKTransform(MagicKey) 'MagicKey = MagicKey * 7 + 3
        tmpInfo.Offset = SetFilePointer(hFile, 0, ByVal 0, FILE_CURRENT) 'get current pointer
        tmpInfo.MagicKey = MagicKey
        ReDim Preserve InfoList(InfoCount) As RGSSAD_SUB_INFO
        InfoList(InfoCount) = tmpInfo
        InfoCount = InfoCount + 1
        SetFilePointer hFile, tmpInfo.FileSize, ByVal 0, FILE_CURRENT
        ret = ReadFile(hFile, tmpInfo.FileNameSize, 4, nBytesRead, ByVal 0)
        If ret = 1 And nBytesRead < 4 Then IsEof = True
    Loop
    InfoCount = InfoCount - 1
    'Extract files
    For i = 0 To InfoCount
        If Not Progress Is Nothing Then
            Progress.Value = Progress.Min + i / InfoCount * (Progress.Max - Progress.Min)
        End If
        s = Replace(Path & "\" & Byt2Str(InfoList(i).FileName), "\\", "\")
        WriteLog hLogFile, "Extracting " & s
        SHCreateDirectoryEx 0, Left$(s, InStrRev(s, "\")), ByVal 0 'create path
        hNewFile = CreateFile(s, FILE_ALL_ACCESS, 0, ByVal 0, CREATE_ALWAYS, 0, 0) 'create output file
        If hFile = -1 Then 'failed
            WriteLog hLogFile, "Create file '" & FileName & "' failed"
            ExtractFile = -1
            GoTo ErrLine
        End If
        leftsize = InfoList(i).FileSize
        MagicKey = InfoList(i).MagicKey
        SetFilePointer hFile, InfoList(i).Offset, ByVal 0, FILE_BEGIN
        Do While leftsize >= 1024
            ReadFile hFile, buff(0), 1024, nBytesRead, ByVal 0
            leftsize = leftsize - 1024
            For j = 0 To 1023 Step 4
                CopyMemory l, buff(j), 4
                l = l Xor MagicKey
                MagicKey = MKTransform(MagicKey) 'MagicKey = MagicKey * 7 + 3
                CopyMemory buff(j), l, 4
            Next
            WriteFile hNewFile, buff(0), 1024, 0, ByVal 0
        Loop
        If leftsize > 0 Then
            ReadFile hFile, buff(0), leftsize, nBytesRead, ByVal 0
            For j = 0 To 1023 Step 4
                CopyMemory l, buff(j), 4
                l = l Xor MagicKey
                MagicKey = MKTransform(MagicKey) 'MagicKey = MagicKey * 7 + 3
                CopyMemory buff(j), l, 4
            Next
            WriteFile hNewFile, buff(0), leftsize, 0, ByVal 0
        End If
        CloseHandle hNewFile
    Next
    WriteLog hLogFile, "Done."
    ExtractFile = 0
ErrLine:
    'Finalize
    CloseHandle hNewFile
    CloseHandle hFile
    CloseHandle hLogFile
End Function

Public Function MagicKeyGuess(FileName As String) As Long
    Dim hFile As Long
    hFile = CreateFile(FileName, FILE_ALL_ACCESS, 0, ByVal 0, OPEN_EXISTING, 0, 0)
    SetFilePointer hFile, 8, ByVal 0, FILE_BEGIN
    ReadFile hFile, MagicKeyGuess, 4, 0, ByVal 0
    CloseHandle hFile
    MagicKeyGuess = MagicKeyGuess Xor 18
End Function

Private Function WriteLog(ByVal hLogFile As Long, LogStr As String, Optional CrLf As Boolean = True) As Long
    Dim s As String
    If CrLf Then
        s = LogStr & vbCrLf
    Else
        s = LogStr
    End If
    WriteLog = WriteFile(hLogFile, ByVal s, LenB(StrConv(s, vbFromUnicode)), 0, ByVal 0)
End Function

Private Function Byt2Str(byt() As Byte) As String
    Dim l As Long
    l = lstrlen(byt(0))
    Byt2Str = String$(l, Chr$(0))
    CopyMemory ByVal Byt2Str, byt(0), l
End Function

Private Function PatchCode() As Long
    '004013DC >    8B4424 04     MOV EAX,DWORD PTR SS:[ESP+4]
    '004013E0      6BC0 07       IMUL EAX,EAX,7
    '004013E3      83C0 03       ADD EAX,3
    '004013E6      C2 1000       RETN 10
    '004013E9      90            NOP
    Const CodeStr = "8B 44 24 04 6B C0 07 83 C0 03 C2 10 00"
    Dim s() As String, i As Long, l As Long
    s = Split(CodeStr, " ")
    l = UBound(s)
    ReDim bCode(l) As Byte
    For i = 0 To l
        bCode(i) = Val("&H" & s(i))
    Next
    'PatchCode = WriteProcessMemory(-1, AddressOf MKTransform, b(0), l + 1, 0)
End Function

Private Function MKTransform(ByVal MagicKey As Long) As Long
    MKTransform = CallWindowProc(VarPtr(bCode(0)), MagicKey, 0, 0, 0)
End Function
