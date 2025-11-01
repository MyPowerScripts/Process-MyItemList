@Echo off

Set MyApp=Process-MyItemList

#"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy ByPass -File "%~dp0%MyApp%.ps1"
"C:\Program Files\PowerShell\7\pwsh.exe" -NoProfile -ExecutionPolicy ByPass -File "%~dp0%MyApp%.ps1"
::Pause

