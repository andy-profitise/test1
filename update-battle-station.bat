@echo off
REM Battle Station Update Script (Batch version)
REM This script pulls the latest changes from git and pushes to Google Apps Script

echo.
echo Updating Battle Station...
echo.

echo Pulling latest changes from GitHub...
git pull

if errorlevel 1 (
    echo.
    echo ERROR: Git pull failed!
    echo Please check your git configuration and try again.
    exit /b 1
)

echo.
echo Git pull successful!
echo.

echo Pushing to Google Apps Script...
clasp push

if errorlevel 1 (
    echo.
    echo ERROR: Clasp push failed!
    echo Please check your clasp configuration and authentication.
    exit /b 1
)

echo.
echo Battle Station updated successfully!
echo Refresh your Google Sheet to see the changes.
echo.
