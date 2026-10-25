@echo off
chcp 65001 >nul
cd /d "%~dp0"
title YouTube Websites Master Hub - 1-Click All-in-One
cls
echo ===============================================================================
echo 🌐 YOUTUBE WEBSITES & AI TOOLS MASTER HUB (1-CLICK ALL-IN-ONE)
echo ===============================================================================
echo.
echo  [Enter]       -^> 🚀 1-Click Auto-Update & Push to GitHub (Default: @setupsaitony)
echo  [Paste URL]   -^> 🔍 Add / Scrape ANY New Channel (e.g. @channel or youtube link)
echo  [V] or [1]    -^> 📖 View Master Report (All Channels Markdown)
echo  [E] or [2]    -^> 📊 Open Master Excel Spreadsheet (.xlsx)
echo  [F] or [3]    -^> 📁 Open All Channels Master Directory Folder
echo.
echo ===============================================================================
set /p TARGET="Choose option or press [Enter] to run 1-Click Auto-Update: "

if /i "%TARGET%"=="V" goto view_md
if /i "%TARGET%"=="1" goto view_md
if /i "%TARGET%"=="E" goto open_excel
if /i "%TARGET%"=="2" goto open_excel
if /i "%TARGET%"=="F" goto open_folder
if /i "%TARGET%"=="3" goto open_folder

echo.
echo [*] Launching Autonomous Scraper & Master Directory Consolidation...
python extract_and_categorize_websites.py "%TARGET%"

echo.
echo ===============================================================================
echo 🎉 1-Click Process Completed! Automatically opening Master Report...
echo ===============================================================================
timeout /t 3 >nul
if exist "📁 All_Channels_Master_Directory\📝 All_Channels_Websites_Master_Categorized_Report.md" (
    start "" "📁 All_Channels_Master_Directory\📝 All_Channels_Websites_Master_Categorized_Report.md"
) else (
    start "" "📝 All_Channels_Websites_Master_Categorized_Report.md"
)
goto end

:view_md
echo [*] Opening Master All-Channels Markdown Report...
if exist "📁 All_Channels_Master_Directory\📝 All_Channels_Websites_Master_Categorized_Report.md" (
    start "" "📁 All_Channels_Master_Directory\📝 All_Channels_Websites_Master_Categorized_Report.md"
) else (
    start "" "📝 All_Channels_Websites_Master_Categorized_Report.md"
)
goto end

:open_excel
echo [*] Opening Master All-Channels Excel Workbook...
if exist "📁 All_Channels_Master_Directory\📊 All_Channels_Websites_Master_List.xlsx" (
    start "" "📁 All_Channels_Master_Directory\📊 All_Channels_Websites_Master_List.xlsx"
) else (
    start "" "📊 All_Channels_Websites_Master_List.xlsx"
)
goto end

:open_folder
echo [*] Opening Master Directory Folder...
if exist "📁 All_Channels_Master_Directory" (
    start "" "📁 All_Channels_Master_Directory"
) else (
    start "" "."
)
goto end

:end
echo.
echo Press any key to exit...
pause >nul
