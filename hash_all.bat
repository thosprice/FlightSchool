@echo off
setlocal enabledelayedexpansion
title Directory SHA-256 Generator

:: --- CONFIGURATION ---
:: Target directory to scan (Defaults to "X-Plane 9" in the current folder)
set "TargetDir=X-Plane 9"
set "OutputFile=hashes.txt"

echo Scanning directory tree: %TargetDir%...
echo Generating %OutputFile%, please wait...

:: Clear the output file if it already exists
if exist "%OutputFile%" del "%OutputFile%"

:: Get the absolute length of the current folder path so we can strip it out later
for %%I in ("%CD%\") do set "RootPath=%%~fI"

:: Loop recursively through every file in the target directory tree
for /R "%TargetDir%" %%F in (*) do (
    :: Get the full absolute windows path of the file
    set "FullPath=%%~fF"
    
    :: Remove the local root path prefix to turn it into a relative path
    set "RelPath=!FullPath:%RootPath%=!"
    
    :: Convert Windows backslashes (\) to Git forward slashes (/) for the manifest layout
    set "RelPath=!RelPath:\=/!"
    
    :: Run certutil to calculate the hash, parsing its text output
    set "FileHash="
    for /f "tokens=* skip=1" %%A in ('certutil -hashfile "%%~fF" SHA256 2^>nul') do (
        if not defined FileHash (
            set "FileHash=%%A"
            :: Strip out any empty space characters from certutil's hash string
            set "FileHash=!FileHash: =!"
        )
    )
    
    :: Write the cleanly formatted path:hash entry straight to the text file
    if defined FileHash (
        echo !RelPath!:!FileHash!>>"%OutputFile%"
    )
)

echo.
echo [SUCCESS] Master hash mapping exported cleanly to %OutputFile%

