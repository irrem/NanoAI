@echo off
echo Building NanoAI...

REM Clean the project
dotnet clean

REM Make sure all files are included
dir /s /b Core\NLU\Handlers\*.cs > filelist.txt

REM Check if handler files exist
echo Checking for handler files:
type filelist.txt
echo.

REM Build the project
dotnet build

echo Done! 