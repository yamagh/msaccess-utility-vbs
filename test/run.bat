@echo off
cls
dir /b test_*.vbs | cscript "%~dp0..\bin\TestRunner.wsf" //nologo //Job:ConsoleTestRunner /stdin+ %*
