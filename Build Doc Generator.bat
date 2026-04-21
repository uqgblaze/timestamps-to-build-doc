@echo off
python "%~dp0build_doc_generator.py"
if errorlevel 1 (
    echo.
    echo Error: failed to launch. Is Python installed?
    pause
)
