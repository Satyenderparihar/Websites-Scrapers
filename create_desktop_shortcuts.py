import os
import sys
import win32com.client

sys.stdout.reconfigure(encoding='utf-8')

desktop_dir = os.path.join(os.environ["USERPROFILE"], "Desktop")
project_dir = r"D:\Antigravity - Agent\🌐 YouTube Websites Extractor Agent"

shell = win32com.client.Dispatch("WScript.Shell")

# 1. Launcher Shortcut (Runs Extractor Agent)
lnk1_path = os.path.join(desktop_dir, "YouTube Websites Extractor.lnk")
sc1 = shell.CreateShortcut(lnk1_path)
sc1.TargetPath = "powershell.exe"
sc1.Arguments = f"-NoExit -NoProfile -ExecutionPolicy Bypass -Command \"& {{ Set-Location -LiteralPath '{project_dir}'; .\\run_agent.bat }}\""
sc1.WorkingDirectory = "D:\\Antigravity - Agent"
sc1.Description = "1-Click Autonomous YouTube Websites Extractor Agent"
sc1.IconLocation = "shell32.dll,13"
sc1.Save()
print(f"✅ Created Desktop Shortcut 1: {lnk1_path}")

# 2. Master Directory Viewer Shortcut (1-Click View All Channels Master Report)
lnk2_path = os.path.join(desktop_dir, "All Channels Websites Master Directory.lnk")
sc2 = shell.CreateShortcut(lnk2_path)
sc2.TargetPath = "powershell.exe"
sc2.Arguments = f"-NoProfile -ExecutionPolicy Bypass -Command \"& {{ Set-Location -LiteralPath '{project_dir}'; .\\view_all_directory.bat }}\""
sc2.WorkingDirectory = "D:\\Antigravity - Agent"
sc2.Description = "1-Click Open All Channels Websites Master Directory (Markdown Report)"
sc2.IconLocation = "shell32.dll,23"
sc2.Save()
print(f"✅ Created Desktop Shortcut 2: {lnk2_path}")
