@echo off
chcp 65001 >nul
echo Starting Filename Auto-Normalizer GUI...
python "%~dp0rename_gui.py" %*
