Option Explicit

On Error Resume Next

Dim objShell, objWindow
Dim Cnt

Set objShell = CreateObject("Shell.Application")

Do
    For Each objWindow In objShell.Windows
        If TypeName(objWindow.document) Like "IShellFolderViewDual*" Then
            objWindow.Quit
            Cnt = Cnt + 1
        End If
    Next
    
    If Cnt = 0 Then Exit Do
    Cnt = 0
Loop

Set objShell = Nothing