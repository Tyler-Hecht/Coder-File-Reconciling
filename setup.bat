@echo off
py clearer.py
if %errorlevel% neq 0 (
	pause
	exit)
py setup1.py
if %errorlevel% neq 0 (
	pause
	exit)
cd DatavyuToSupercoder
java -jar DatavyuToSupercoder.jar
cd ..
py setup2.py
if %errorlevel% neq 0 (
	pause
	exit)
pause