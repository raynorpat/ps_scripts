@echo off
echo joining domain

powershell -ExecutionPolicy bypass -File .\join-domain.ps1

pause