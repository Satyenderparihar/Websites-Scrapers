$wsh = New-Object -ComObject WScript.Shell
$desktop = [System.Environment]::GetFolderPath('Desktop')
$targetBat = "D:\Antigravity - Agent\🌐 YouTube Websites Extractor Agent\run_agent.bat"
$workingDir = "D:\Antigravity - Agent\🌐 YouTube Websites Extractor Agent"

# Shortcut 1: Clean ASCII name (guaranteed compatibility across all Windows versions)
$sc1Path = Join-Path $desktop "YouTube Websites Extractor.lnk"
$sc1 = $wsh.CreateShortcut($sc1Path)
$sc1.TargetPath = $targetBat
$sc1.WorkingDirectory = $workingDir
$sc1.Description = "YouTube Websites Extractor & Categorizer Agent"
$sc1.IconLocation = "shell32.dll,13"  # Globe icon
$sc1.Save()
Write-Output "Created $sc1Path"

# Shortcut 2: Update existing emoji named shortcut if needed
$sc2Path = Join-Path $desktop "🌐 YouTube Websites Extractor.lnk"
$sc2 = $wsh.CreateShortcut($sc2Path)
$sc2.TargetPath = $targetBat
$sc2.WorkingDirectory = $workingDir
$sc2.Description = "YouTube Websites Extractor & Categorizer Agent"
$sc2.IconLocation = "shell32.dll,13"  # Globe icon
$sc2.Save()
Write-Output "Updated $sc2Path"
