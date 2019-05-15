@echo off
powershell Invoke-ScriptAnalyzer ppt2pdf.ps1
powershell -ExecutionPolicy Bypass -command "&ps2exe.ps1 -inputFile ppt2pdf.ps1 -outputFile ppt2pdf.exe"
del ppt2pdf.exe.config
