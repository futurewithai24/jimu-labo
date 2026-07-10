@echo off
powershell -Command "Write-Host ''; Write-Host '===================================='; Write-Host '  じむラボ - 社内共有URL'; Write-Host ''; Write-Host '  https://jimu-labo.streamlit.app/'; Write-Host ''; Write-Host '  このURLを社内の人に共有してください'; Write-Host '====================================';"
powershell -Command "Start-Process 'https://jimu-labo.streamlit.app/'"
pause
