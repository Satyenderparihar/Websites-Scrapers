$WshShell = New-Object -ComObject WScript.Shell
$DesktopPath = [System.Environment]::GetFolderPath('Desktop')
$ShortcutPath = Join-Path $DesktopPath '🌐 YouTube Websites Extractor.lnk'
$Shortcut = $WshShell.CreateShortcut($ShortcutPath)
$Shortcut.TargetPath = 'D:\Antigravity - Agent\🌐 YouTube Websites Extractor Agent\⚡ Run YouTube Extractor Agent.bat'
$Shortcut.WorkingDirectory = 'D:\Antigravity - Agent\🌐 YouTube Websites Extractor Agent'
$Shortcut.Description = 'Launch YouTube Websites Extractor & Categorizer Agent'
$Shortcut.Save()
Write-Output "Shortcut successfully created at $ShortcutPath"
