# Battle Station - Google Apps Script Setup

This repository contains the Battle Station vendor review dashboard for Google Sheets.

## 📦 Project Structure

```
.
├── BattleStation.gs              # Main Apps Script code
├── appsscript.json               # Apps Script manifest
├── update-battle-station.bat     # Quick update script for Windows (batch)
├── update-battle-station.ps1     # Quick update script for Windows (PowerShell)
├── .clasp.json.template          # Template for clasp configuration
├── .gitattributes                # Git attributes for proper file handling
├── .gitignore                    # Git ignore file
└── README.md                     # This file
```

## 🚀 Setup Instructions

### Prerequisites

- Node.js and npm installed ✅ (Already installed)
- clasp installed globally ✅ (Already installed)
- A Google account with access to Google Sheets and Apps Script

### Step 1: Authenticate with Google

```bash
clasp login
```

This will open your browser to authenticate with your Google account. Sign in and grant permissions.

### Step 2: Choose Your Setup Method

You have two options:

#### Option A: Link to Existing Apps Script Project (Recommended if you already have one)

1. Open your Google Sheet with the Battle Station
2. Go to **Extensions > Apps Script**
3. In the Apps Script editor, click on **Project Settings** (gear icon)
4. Copy the **Script ID**
5. Create `.clasp.json` from the template:

```bash
cp .clasp.json.template .clasp.json
```

6. Edit `.clasp.json` and replace `YOUR_SCRIPT_ID_HERE` with your actual Script ID:

```json
{
  "scriptId": "YOUR_ACTUAL_SCRIPT_ID",
  "rootDir": "/home/user/test1"
}
```

#### Option B: Create a New Apps Script Project

1. Create a new standalone Apps Script project:

```bash
clasp create --title "Battle Station" --type standalone
```

Or create one bound to a specific spreadsheet:

```bash
clasp create --title "Battle Station" --type sheets --parentId "YOUR_SPREADSHEET_ID"
```

This will automatically create the `.clasp.json` file for you.

### Step 3: Push Your Code to Google

```bash
clasp push
```

This uploads `BattleStation.gs` and `appsscript.json` to Google Apps Script.

**Note:** The first time you push, clasp may ask you to confirm overwriting files. Type `yes` to continue.

### Step 4: Open Your Project

```bash
clasp open
```

This opens your Apps Script project in your browser.

## 🔄 Development Workflow

### Quick Update (Recommended for Windows)

Use a helper script to pull from git and push to Google in one command:

**Option 1: Batch File (Most Compatible)**
```cmd
update-battle-station.bat
```

**Option 2: PowerShell (Colorful Output)**
```powershell
./update-battle-station.ps1
```

Both scripts will:
1. Pull the latest changes from GitHub
2. Push them to Google Apps Script
3. Show you status messages along the way

Note: If you get encoding errors with the PowerShell script, use the batch file instead.

### Making Changes Manually

1. Edit `BattleStation.gs` locally
2. Push changes to Google Apps Script:

```bash
clasp push
```

3. Test in your Google Sheet

### Pulling Changes

If you make changes directly in the Apps Script editor:

```bash
clasp pull
```

This downloads the latest version from Google to your local files.

### Watching for Changes

For continuous development:

```bash
clasp push --watch
```

This automatically pushes changes whenever you save a file.

## 📝 Git Workflow

The `.clasp.json` file is gitignored because it contains user-specific configuration. Each developer should:

1. Clone the repository
2. Run `clasp login`
3. Create their own `.clasp.json` using the template
4. Link to their own Apps Script project

## 🛠️ Useful clasp Commands

| Command | Description |
|---------|-------------|
| `clasp login` | Authenticate with Google |
| `clasp push` | Upload local files to Google |
| `clasp pull` | Download files from Google |
| `clasp open` | Open the script in browser |
| `clasp logs` | View execution logs |
| `clasp deployments` | List deployments |
| `clasp version` | Create a new version |

## 📋 OAuth Scopes

This script requires the following permissions (defined in `appsscript.json`):

- `spreadsheets.currentonly` - Read/write access to the current spreadsheet
- `gmail.readonly` - Read Gmail messages (for email search functionality)
- `script.external_request` - Make external API calls (for monday.com integration)

## 🔧 Troubleshooting

### "User has not enabled the Apps Script API"

1. Go to https://script.google.com/home/usersettings
2. Enable **Google Apps Script API**

### Authentication Issues

```bash
clasp logout
clasp login
```

### Push/Pull Conflicts

```bash
clasp pull  # Get latest from Google
# Resolve conflicts
clasp push  # Push back to Google
```

## 📚 Resources

- [clasp Documentation](https://github.com/google/clasp)
- [Apps Script Documentation](https://developers.google.com/apps-script)
- [Apps Script API Reference](https://developers.google.com/apps-script/reference)
