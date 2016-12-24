Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function VirtualAllocEx Lib "kernel32.dll" (ByVal hProcess As Long, lpAddress As Any, ByRef dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Public Declare Function VirtualFreeEx Lib "kernel32.dll" (ByVal hProcess As Long, lpAddress As Any, ByRef dwSize As Long, ByVal dwFreeType As Long) As Long
Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As Long, lpThreadAttributes As Any, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function lstrcmp Lib "kernel32" Alias "lstrcmpA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function GetProcessId Lib "kernel32" (ByVal hProcess As Long) As Long
Public Type POINTAPI
    x As Long
    y As Long
End Type
Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Const MEM_RESERVE = &H2000
Public Const MEM_COMMIT = &H1000
Public Const MEM_RELEASE = &H8000
Public Const PAGE_READWRITE = &H4
Public Const DONT_RESOLVE_DLL_REFERENCES = &H1

Function GetProcessNameByProcessId(ByVal pid As Long) As String
    Dim szBuf(1 To 250) As Long
    Dim Ret As Long
    Dim szPathName As String
    Dim nSize As Long
    Dim hProcess As Long
    hProcess = OpenProcess(&H400 Or &H10, 0, pid)
    If hProcess <> 0 Then
        Ret = EnumProcessModules(hProcess, szBuf(1), 250, pid)
        If Ret <> 0 Then
            szPathName = Space(260)
            nSize = 500
            Ret = GetModuleFileNameExA(hProcess, szBuf(1), szPathName, nSize)
            GetProcessNameByProcessId = szPathName
        End If
    End If
    Ret = CloseHandle(hProcess)
End Function

Function GetFileName(Path As String)
    GetFileName = Right(Path, InStr(StrReverse(Path), "\") - 1)
End Function

Function GetRGSSEvalAddress(ByVal hProcess As Long, sDLLName As String) As Long
    Dim szBuf(1 To 250) As Long, Ret As Long, szPathName As String, nSize As Long, I As Long, hDLL As Long, pid As Long
    pid = GetProcessId(hProcess)
    Ret = EnumProcessModules(hProcess, szBuf(1), 250, pid)
    If Ret <> 0 Then
        For I = 1 To 250
            szPathName = Space(260)
            nSize = 500
            Ret = GetModuleFileNameExA(hProcess, szBuf(I), szPathName, nSize)
            If lstrcmp(GetFileName(szPathName), sDLLName) = 0 Then
                hDLL = LoadLibraryEx(szPathName, 0, DONT_RESOLVE_DLL_REFERENCES)
                If hDLL = 0 Then Exit Function
                GetRGSSEvalAddress = szBuf(I) + GetProcAddress(hDLL, "RGSSEval") - hDLL
                FreeLibrary hDLL
                Exit Function
            End If
        Next
    End If
End Function

Function RemoteCall(ByVal hProcess As Long, ByVal pfnRemoteFunc As Long, ByVal pLocalParam As String, Optional ByVal bWait As Boolean = False) As Long
    Dim pRemoteParam As Long, hThread As Long, dwExitCode As Long, nParamSize As Long
    nParamSize = 1 + LenB(StrConv(pLocalParam, vbFromUnicode))
    pRemoteParam = VirtualAllocEx(hProcess, ByVal 0&, nParamSize, MEM_RESERVE Or MEM_COMMIT, PAGE_READWRITE)
    If pRemoteParam = 0 Then GoTo ErrLine
    If WriteProcessMemory(hProcess, ByVal pRemoteParam, ByVal pLocalParam, nParamSize, 0) = 0 Then GoTo ErrLine
    hThread = CreateRemoteThread(hProcess, ByVal 0, 0, ByVal pfnRemoteFunc, ByVal pRemoteParam, 0, ByVal 0)
    If hThread = 0 Then GoTo ErrLine
    RemoteCall = hThread
    If bWait Then
        WaitForSingleObject hThread, -1
        GetExitCodeThread hThread, dwExitCode
        RemoteCall = dwExitCode
        CloseHandle hThread
    End If
ErrLine:
    If pRemoteParam Then VirtualFreeEx hProcess, pRemoteParam, nParamSize, MEM_RELEASE
End Function
