@echo off
REM Refreshes every FT cheat sheet / report. Run by the "Wofford FT Cheat Sheet Update" scheduled task,
REM or double-click to update right now. Output of the last run goes to update_log.txt.
cd /d "%~dp0"
echo ==== %date% %time% ==== > update_log.txt
py ft_cheat_sheet.py >> update_log.txt 2>&1
