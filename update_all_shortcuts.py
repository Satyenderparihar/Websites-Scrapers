import os
import sys
sys.stdout.reconfigure(encoding='utf-8')
import win32com.client

wsh = win32com.client.Dispatch('WScript.Shell')
desktop = os.path.expanduser('~\\Desktop')

cmd_args = (
    '-NoExit -NoProfile -ExecutionPolicy Bypass -Command '
    '"& { $d = (Get-ChildItem \'D:\\Antigravity - Agent\' -Directory | '
    'Where-Object { $_.Name -like \'*YouTube*\' } | Select-Object -First 1).FullName; '
    'Set-Location -LiteralPath $d; .\\run_agent.bat }"'
)

for fn in os.listdir(desktop):
    if 'YouTube' in fn and fn.endswith('.lnk'):
        fp = os.path.join(desktop, fn)
        try:
            sc = wsh.CreateShortcut(fp)
            sc.TargetPath = 'powershell.exe'
            sc.Arguments = cmd_args
            sc.WorkingDirectory = r'D:\Antigravity - Agent'
            sc.Description = 'YouTube Websites Extractor & Categorizer Agent'
            sc.IconLocation = 'shell32.dll,13'
            sc.Save()
            print(f'Successfully updated: {fn}')
        except Exception as e:
            print(f'Error updating {fn}: {e}')

# Verify each shortcut
print('\n--- VERIFICATION ---')
for fn in os.listdir(desktop):
    if 'YouTube' in fn and fn.endswith('.lnk'):
        fp = os.path.join(desktop, fn)
        sc = wsh.CreateShortcut(fp)
        print(f'Shortcut: {fn}')
        print(f'  Target: {sc.TargetPath}')
        print(f'  Args: {sc.Arguments[:60]}...')
        print(f'  WorkingDir: {sc.WorkingDirectory}')
