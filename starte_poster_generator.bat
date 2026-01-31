@echo off
set "SCRIPT_DIR=%~dp0"
python "%SCRIPT_DIR%poster_generator.py" %*
pause
