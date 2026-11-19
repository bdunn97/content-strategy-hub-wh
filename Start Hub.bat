@echo off
rem Double-click to start the Influencer Content Hub at http://localhost:4324
start "" http://localhost:4324
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0serve.ps1"
