@echo off
powershell -Command "Start-Process powershell -ArgumentList '-ExecutionPolicy Bypass -File \"%~dp0DebloatGUI.ps1\"' -Verb RunAs" 