Imports System
Imports System.Runtime.InteropServices
Imports Microsoft.Win32
Imports System.Diagnostics
Imports System.Collections.Generic

Public Module MyApplication

    Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As IntPtr, ByVal lpString As String) As Boolean

    Sub Main()
        On Error Resume Next
        Dim CmdLine As String
        Dim Args() As String

        CmdLine = Command()
        Args = ParseCommandLine(CmdLine)

        If Args.Length >= 3 Then
            If Args(0).ToLower() = "/pid" Then
                Dim targetPid As Integer
                If Integer.TryParse(Args(1), targetPid) Then
                    ChangeTitleByPid(targetPid, Args(2))
                Else
                    Console.WriteLine("Invalid process ID.")
                End If
            ElseIf Args(0).ToLower() = "/windowtorename" Then
                Dim oldName As String = Args(1)
                Dim newName As String = Args(2)
                ChangeTitleByOldName(oldName, newName)
            Else
                ShowUsage()
            End If
        Else
            ShowUsage()
        End If
    End Sub

    Private Sub ShowUsage()
		MsgBox("Usage:" & vbCrLf & vbCrLf & "settext.exe /WindowToRename ""Old name"" ""New Name""" & vbCrLf & vbCrLf & "settext.exe /PID ProcessID ""New Name""")
        Console.WriteLine("Usage:")
        Console.WriteLine("To change the title by process ID: ChangeTitle /pid [Process ID] [New Title]")
        Console.WriteLine("To change the title by old name: ChangeTitle /oldname [Old Name] [New Name]")
    End Sub

    Private Sub ChangeTitleByPid(pid As Integer, newTitle As String)
        Dim processes() As Process = Process.GetProcesses()
        For Each proc As Process In processes
            If proc.Id = pid Then
                SetWindowText(proc.MainWindowHandle, newTitle)
                Exit For
            End If
        Next
    End Sub

    Private Sub ChangeTitleByOldName(oldName As String, newTitle As String)
        Dim processes() As Process = Process.GetProcesses()
        For Each proc As Process In processes
            If proc.MainWindowTitle = oldName Then
                SetWindowText(proc.MainWindowHandle, newTitle)
                Exit For
            End If
        Next
    End Sub

    Private Function ParseCommandLine(commandLine As String) As String()
        Dim args As New List(Of String)
        Dim inQuotes = False
        Dim currentArg As String = ""
        For Each c As Char In commandLine
            If c = """"c Then
                inQuotes = Not inQuotes
            ElseIf c = " "c AndAlso Not inQuotes Then
                args.Add(currentArg)
                currentArg = ""
            Else
                currentArg &= c
            End If
        Next
        args.Add(currentArg)
        Return args.ToArray()
    End Function
End Module
