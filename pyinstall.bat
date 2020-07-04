@echo off
rd /S /Q dist build
pyinstaller -F -i logo.ico %1 -w
Pause