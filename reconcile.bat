@echo off
py reconcile.py
if %errorlevel% neq 0 (
	pause
	exit)
pause
