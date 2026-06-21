# Release Instructions

This repository is configured to publish automatically to the npm registry via GitHub Actions whenever a new GitHub Release is created.

## Step-by-Step Release Guide

### 1. Verify Build
```powershell
npm run build; npm run lint
```

### 2. Bump Version & Create Tag
```bash
npm version patch    # Bugfix
npm version minor    # New feature
npm version major    # Breaking change
```

### 3. Push to GitHub
```bash
git push origin main --follow-tags
```

### 4. Create GitHub Release
Go to GitHub → Releases → Create from the new tag → Actions auto-publishes.

## Manual Publish (Fallback)
See `npm_publishing_guide.txt` in workspace root.
