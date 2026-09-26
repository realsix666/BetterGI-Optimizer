Option Explicit

Dim fso, shell, scriptPath, betterGIPath
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
scriptPath = WScript.ScriptFullName


betterGIPath = FindBetterGI()

If betterGIPath = "" Then
    MsgBox "找不到 BetterGI，请确认已安装！", 16, "错误"
    WScript.Quit
End If

'
CreateShortcutWithIcon scriptPath, betterGIPath


shell.Run """" & betterGIPath & """", 0, False
WScript.Sleep 3000


Dim dollar, bindCmd
dollar = Chr(36)
bindCmd = "powershell -WindowStyle Hidden -Command """ & _
          dollar & "p = Get-Process 'BetterGI' -EA SilentlyContinue; " & _
          "if(" & dollar & "p){ " & _
          "  " & dollar & "cores = [Environment]::ProcessorCount; " & _
          "  " & dollar & "mask = [Math]::Pow(2," & dollar & "cores) - [Math]::Pow(2,[Math]::Floor(" & dollar & "cores/2)); " & _
          "  " & dollar & "p.ProcessorAffinity = [long]" & dollar & "mask " & _
          "}"""
shell.Run bindCmd, 0, True



Function FindBetterGI()
    Dim candidates, i, sk, val
    Dim subKeys, wmi, procs, proc, found

    
    candidates = Array( _
        shell.ExpandEnvironmentStrings("%ProgramFiles%\BetterGI\BetterGI.exe"), _
        shell.ExpandEnvironmentStrings("%ProgramFiles(x86)%\BetterGI\BetterGI.exe"), _
        shell.ExpandEnvironmentStrings("%LocalAppData%\BetterGI\BetterGI.exe"), _
        shell.ExpandEnvironmentStrings("%AppData%\BetterGI\BetterGI.exe"), _
        shell.ExpandEnvironmentStrings("%ProgramData%\BetterGI\BetterGI.exe"), _
        "C:\BetterGI\BetterGI.exe", _
        "D:\BetterGI\BetterGI.exe", _
        "D:\Program Files\BetterGI\BetterGI.exe", _
        "E:\BetterGI\BetterGI.exe" _
    )
    For i = 0 To UBound(candidates)
        If fso.FileExists(candidates(i)) Then
            FindBetterGI = candidates(i)
            Exit Function
        End If
    Next

    
    On Error Resume Next
    subKeys = Array( _
        "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\BetterGI\", _
        "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\BetterGI\", _
        "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\BetterGI\" _
    )
    For Each sk In subKeys
        Err.Clear
        val = shell.RegRead(sk & "InstallLocation")
        If Err.Number = 0 And Len(val) > 0 Then
            If fso.FileExists(val & "\BetterGI.exe") Then
                FindBetterGI = val & "\BetterGI.exe"
                Exit Function
            End If
        End If

        Err.Clear
        val = shell.RegRead(sk & "DisplayIcon")
        If Err.Number = 0 And Len(val) > 0 Then
            val = Replace(val, """", "")
            val = Replace(val, ",0", "")
            If fso.FileExists(val) Then
                FindBetterGI = val
                Exit Function
            End If
        End If
    Next
    On Error GoTo 0

    
    On Error Resume Next
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set procs = wmi.ExecQuery("SELECT ExecutablePath FROM Win32_Process WHERE Name='BetterGI.exe'")
    For Each proc In procs
        If Not IsNull(proc.ExecutablePath) Then
            FindBetterGI = proc.ExecutablePath
            Exit Function
        End If
    Next
    On Error GoTo 0

    
    found = SearchStartMenu()
    If found <> "" Then
        FindBetterGI = found
        Exit Function
    End If

    FindBetterGI = ""
End Function



Function SearchStartMenu()
    Dim dirs, d, lnkPath, sc
    SearchStartMenu = ""
    dirs = Array( _
        shell.ExpandEnvironmentStrings("%AppData%\Microsoft\Windows\Start Menu\Programs"), _
        shell.ExpandEnvironmentStrings("%ProgramData%\Microsoft\Windows\Start Menu\Programs") _
    )
    On Error Resume Next
    For Each d In dirs
        lnkPath = RecursiveFindLnk(d, "bettergi", 0)
        If lnkPath <> "" Then
            Set sc = shell.CreateShortcut(lnkPath)
            If fso.FileExists(sc.TargetPath) Then
                SearchStartMenu = sc.TargetPath
                Exit Function
            End If
        End If
    Next
    On Error GoTo 0
End Function



Function RecursiveFindLnk(folderPath, keyword, depth)
    Dim folder, sub, f, r
    RecursiveFindLnk = ""
    If depth > 4 Then Exit Function
    If Not fso.FolderExists(folderPath) Then Exit Function

    On Error Resume Next
    Set folder = fso.GetFolder(folderPath)
    If Err.Number <> 0 Then Exit Function

    For Each f In folder.Files
        If LCase(fso.GetExtensionName(f.Name)) = "lnk" Then
            If InStr(LCase(f.Name), keyword) > 0 Then
                RecursiveFindLnk = f.Path
                Exit Function
            End If
        End If
    Next

    For Each sub In folder.SubFolders
        r = RecursiveFindLnk(sub.Path, keyword, depth + 1)
        If r <> "" Then
            RecursiveFindLnk = r
            Exit Function
        End If
    Next
    On Error GoTo 0
End Function



Sub CreateShortcutWithIcon(vbsPath, iconExePath)
    Dim lnkPath, shortcut
    lnkPath = fso.GetParentFolderName(vbsPath) & "\" & fso.GetBaseName(vbsPath) & ".lnk"

    On Error Resume Next
    Set shortcut = shell.CreateShortcut(lnkPath)
    shortcut.TargetPath = "wscript.exe"
    shortcut.Arguments = """" & vbsPath & """"
    shortcut.WorkingDirectory = fso.GetParentFolderName(vbsPath)
    shortcut.IconLocation = iconExePath & ",0"
    shortcut.Description = "启动 BetterGI"
    shortcut.Save
    On Error GoTo 0
End Sub
