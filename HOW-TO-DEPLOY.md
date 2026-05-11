# BwB Word Add-in — Deployment Guide

## Step 1: Create the GitHub repo
1. Go to github.com and create a new repo called `bwb-word-addin`
2. Make it **Public** (required for GitHub Pages free hosting)

## Step 2: Upload these files
Upload everything in this `dist/` folder to the repo root.
The simplest way: drag-and-drop all files into the GitHub web interface.

## Step 3: Enable GitHub Pages
1. Go to repo Settings → Pages
2. Source: **Deploy from a branch**
3. Branch: **main**, folder: **/ (root)**
4. Save. GitHub will give you a URL like:
   `https://YOUR-USERNAME.github.io/bwb-word-addin/`

## Step 4: Update the manifest
Open `manifest.xml` and replace every instance of:
  `YOUR-GITHUB-USERNAME`
with your actual GitHub username. Then re-upload the updated manifest.xml.

## Step 5: Sideload in Word (per user, once)
### Windows:
1. Copy `manifest.xml` to a shared network folder or local folder
2. In Word: File → Options → Trust Center → Trust Center Settings
   → Trusted Add-in Catalogs → add the folder path → restart Word
3. Insert → My Add-ins → Shared Folder → BwB Document Assistant

### Mac:
1. Copy `manifest.xml` to:
   `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/`
2. Restart Word → Insert → My Add-ins → BwB Document Assistant

### Word Online (any browser):
1. Insert → Add-ins → Upload My Add-in → upload manifest.xml

## Step 6: IT deployment (optional — suppresses per-user setup)
If you have M365 Admin access (ask Lesley):
- Admin Centre → Settings → Integrated Apps → Upload custom apps
- Upload manifest.xml — it deploys to all BwB users automatically
