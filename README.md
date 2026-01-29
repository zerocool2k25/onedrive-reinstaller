# OneDrive Reinstaller

Restores OneDrive functionality after it has been completely removed by system optimization tools.

## Download

Download the latest version from the [Releases](../../releases) page.

## What it does

This tool reverses common OneDrive removal techniques by:

1. **Removing Group Policy blocks** - Clears `DisableFileSyncNGSC` and `DisableFileSync` policies
2. **Restoring Explorer integration** - Restores the OneDrive CLSID so it appears in File Explorer
3. **Restoring Desktop namespace** - Makes OneDrive appear in the File Explorer sidebar
4. **Restoring context menus** - Brings back the right-click OneDrive options
5. **Creating the OneDrive folder** - Recreates the user's OneDrive folder if deleted
6. **Installing OneDrive** - Downloads and installs OneDrive fresh from Microsoft

## Usage

1. Download `OneDrive_Reinstaller.exe` from the Releases page
2. Right-click the file and select **"Run as administrator"**
3. Follow the on-screen prompts
4. **Restart your computer** after completion
5. Sign in to OneDrive with your Microsoft account

## Requirements

- Windows 10 or Windows 11
- Administrator privileges
- Internet connection (to download OneDrive)

## If OneDrive still won't install

If the automatic installation fails:

1. Download OneDrive manually: https://onedrive.live.com/about/download
2. Or search "OneDrive" in the Microsoft Store

## Building from source

```bash
pip install pyinstaller
pyinstaller --onefile --name "OneDrive_Reinstaller" --uac-admin --console onedrive_reinstaller.py
```

The executable will be in the `dist` folder.

## License

MIT License
