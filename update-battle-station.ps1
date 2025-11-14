# Battle Station Update Script
# This script pulls the latest changes from git and pushes to Google Apps Script

Write-Host "🚀 Updating Battle Station..." -ForegroundColor Cyan
Write-Host ""

# Step 1: Pull latest changes from git
Write-Host "📥 Pulling latest changes from GitHub..." -ForegroundColor Yellow
git pull

if ($LASTEXITCODE -ne 0) {
    Write-Host "❌ Git pull failed!" -ForegroundColor Red
    Write-Host "Please check your git configuration and try again." -ForegroundColor Red
    exit 1
}

Write-Host "✅ Git pull successful!" -ForegroundColor Green
Write-Host ""

# Step 2: Push to Google Apps Script
Write-Host "☁️  Pushing to Google Apps Script..." -ForegroundColor Yellow
clasp push

if ($LASTEXITCODE -ne 0) {
    Write-Host "❌ Clasp push failed!" -ForegroundColor Red
    Write-Host "Please check your clasp configuration and authentication." -ForegroundColor Red
    exit 1
}

Write-Host "✅ Clasp push successful!" -ForegroundColor Green
Write-Host ""
Write-Host "🎉 Battle Station updated successfully!" -ForegroundColor Cyan
Write-Host "Refresh your Google Sheet to see the changes." -ForegroundColor White
