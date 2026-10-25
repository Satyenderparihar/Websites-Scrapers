import os
import sys
import win32com.client

sys.stdout.reconfigure(encoding='utf-8')

desktop_dir = os.path.join(os.environ["USERPROFILE"], "Desktop")
project_dir = r"D:\Antigravity - Agent\🌐 YouTube Websites Extractor Agent"

shell = win32com.client.Dispatch("WScript.Shell")

# The Single Unified 1-Click Master Hub Launcher
unified_target = "powershell.exe"
unified_args = f"-NoExit -NoProfile -ExecutionPolicy Bypass -Command \"& {{ Set-Location -LiteralPath '{project_dir}'; .\\run_agent.bat }}\""

# Primary Master Hub Shortcut
hub_lnk = os.path.join(desktop_dir, "YouTube Websites Master Hub.lnk")
sc_hub = shell.CreateShortcut(hub_lnk)
sc_hub.TargetPath = unified_target
sc_hub.Arguments = unified_args
sc_hub.WorkingDirectory = "D:\\Antigravity - Agent"
sc_hub.Description = "1-Click YouTube Websites Master Hub (Scrape, Consolidate & View Master Directory)"
sc_hub.IconLocation = "shell32.dll,13"
sc_hub.Save()
print(f"✅ Created Primary Master Hub Shortcut: {hub_lnk}")

# Update existing shortcuts so user can never click the wrong one
for alias_name in ["YouTube Websites Extractor.lnk", "All Channels Websites Master Directory.lnk"]:
    alias_path = os.path.join(desktop_dir, alias_name)
    sc = shell.CreateShortcut(alias_path)
    sc.TargetPath = unified_target
    sc.Arguments = unified_args
    sc.WorkingDirectory = "D:\\Antigravity - Agent"
    sc.Description = "1-Click YouTube Websites Master Hub"
    sc.IconLocation = "shell32.dll,13"
    sc.Save()
    print(f"✅ Synced Shortcut: {alias_path}")
