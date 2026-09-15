@echo off
set "target_dir=%~1"

:: If no directory is dragged-and-dropped or typed, use the current folder
if "%target_dir%"=="" set "target_dir=%cd%"

echo Scanning directory tree: %target_dir%
echo Please wait...

:: Loop through all files in the directory and all subfolders
(for /R "%target_dir%" %%i in (*) do (
    echo File: %%i 
    certutil -hashfile "%%i" SHA256 
)) > "%target_dir%\tree_hashes.txt"

echo Done! Results saved to %target_dir%\tree_hashes.txt
pause
