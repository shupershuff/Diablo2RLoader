Imports System
Imports System.Runtime.InteropServices
Imports Microsoft.Win32

Public Module MyApplication  

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Sub Main()
    On Error Resume Next
    Dim CmdLine As String
    Dim Ret as Long
    Dim A() as String
    Dim hwindows as long

    CmdLine = Command()
    If Left(CmdLine, 2) = "/?" Then
        MsgBox("Usage:" & vbCrLf & vbCrLf & "ChangeTitleBar Oldname NewName")
    Else
        A = Split(CmdLine, Chr(34), -1, vbBinaryCompare)
        hwindows = FindWindow(vbNullString, A(1))
        Ret = SetWindowText(hwindows, A(3))

    End If
End Sub
End Module
