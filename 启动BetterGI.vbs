betterGIPath = "BetterGI路径"

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

If Not fso.FileExists(betterGIPath) Then
    MsgBox "找不到 BetterGI，请编辑脚本修改路径！", 16, "错误"
    WScript.Quit
End If

shell.Run """" & betterGIPath & """", 0, False

WScript.Sleep 3000

dollar = Chr(36)
bindCmd = "powershell -WindowStyle Hidden -Command """ & _
          dollar & "p = Get-Process 'BetterGI' -EA SilentlyContinue; " & _
          "if(" & dollar & "p){ " & _
          "  " & dollar & "cores = [Environment]::ProcessorCount; " & _
          "  " & dollar & "mask = [Math]::Pow(2," & dollar & "cores) - [Math]::Pow(2,[Math]::Floor(" & dollar & "cores/2)); " & _
          "  " & dollar & "p.ProcessorAffinity = [long]" & dollar & "mask " & _
          "}"""
shell.Run bindCmd, 0, True
