$wsh = New-Object -ComObject WScript.Shell
$desktop = [System.Environment]::GetFolderPath('Desktop')

$scPath = Join-Path $desktop "YouTube Websites Extractor.lnk"
$sc = $wsh.CreateShortcut($scPath)
$sc.TargetPath = "powershell.exe"
$sc.Arguments = "-NoExit -NoProfile -ExecutionPolicy Bypass -Command `"& { `$d = (Get-ChildItem 'D:\Antigravity - Agent' -Directory | Where-Object { `$_.Name -like '*YouTube*' } | Select-Object -First 1).FullName; Set-Location -LiteralPath `$d; .\run_agent.bat }`""
$sc.WorkingDirectory = "D:\Antigravity - Agent"
$sc.Description = "YouTube Websites Extractor & Categorizer Agent"
$sc.IconLocation = "shell32.dll,13"
$sc.Save()

Write-Output "Successfully updated $scPath"

# Also update the emoji named one if it exists
$scPath2 = Join-Path $desktop "🌐 YouTube Websites Extractor.lnk"
$sc2 = $wsh.CreateShortcut($scPath2)
$sc2.TargetPath = "powershell.exe"
$sc2.Arguments = "-NoExit -NoProfile -ExecutionPolicy Bypass -Command `"& { `$d = (Get-ChildItem 'D:\Antigravity - Agent' -Directory | Where-Object { `$_.Name -like '*YouTube*' } | Select-Object -First 1).FullName; Set-Location -LiteralPath `$d; .\run_agent.bat }`""
$sc2.WorkingDirectory = "D:\Antigravity - Agent"
$sc2.Description = "YouTube Websites Extractor & Categorizer Agent"
$sc2.IconLocation = "shell32.dll,13"
$sc2.Save()

Write-Output "Successfully updated $scPath2"
