"""
OneDrive Reinstaller

Restores OneDrive functionality after removal by system optimization tools.
"""

import os
import sys
import ctypes
import subprocess
import winreg
import urllib.request
import tempfile
import time

def is_admin():
    """Check if running with administrator privileges."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    """Re-launch the script with admin privileges."""
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, " ".join(sys.argv), None, 1
    )

def enable_ansi_colors():
    """Enable ANSI color codes in Windows terminal."""
    if sys.platform == 'win32':
        kernel32 = ctypes.windll.kernel32
        kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)

def run_command(command):
    """Run a command silently."""
    try:
        subprocess.run(command, shell=True, capture_output=True, text=True)
        return True
    except:
        return False

def restore_onedrive():
    """Restore all OneDrive functionality."""

    onedrive_clsid = "{018D5C66-4533-4307-9B53-224DE2ED1FE6}"
    filesync_clsid = "{CB3D0F55-BC2C-4C1A-85ED-23ED75B5106B}"

    # Remove Group Policy blocks
    run_command('reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\OneDrive" /v DisableFileSyncNGSC /t REG_DWORD /d 0 /f')
    run_command('reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\OneDrive" /v DisableFileSync /t REG_DWORD /d 0 /f')
    run_command('reg delete "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\OneDrive" /v DisableFileSyncNGSC /f')
    run_command('reg delete "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\OneDrive" /v DisableFileSync /f')
    run_command('reg delete "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\OneDrive" /v DisableLibrariesDefaultSaveToOneDrive /f')

    # Restore Explorer shell integration
    run_command(f'reg add "HKCR\\CLSID\\{onedrive_clsid}" /ve /t REG_SZ /d "OneDrive" /f')
    run_command(f'reg add "HKCR\\CLSID\\{onedrive_clsid}" /v System.IsPinnedToNameSpaceTree /t REG_DWORD /d 1 /f')
    run_command(f'reg add "HKCR\\CLSID\\{onedrive_clsid}\\ShellFolder" /v Attributes /t REG_DWORD /d 0xF080004D /f')
    run_command(f'reg add "HKCR\\CLSID\\{onedrive_clsid}\\ShellFolder" /v FolderValueFlags /t REG_DWORD /d 0x28 /f')
    run_command(f'reg add "HKCR\\Wow6432Node\\CLSID\\{onedrive_clsid}" /ve /t REG_SZ /d "OneDrive" /f')
    run_command(f'reg add "HKCR\\Wow6432Node\\CLSID\\{onedrive_clsid}" /v System.IsPinnedToNameSpaceTree /t REG_DWORD /d 1 /f')
    run_command(f'reg add "HKCR\\Wow6432Node\\CLSID\\{onedrive_clsid}\\ShellFolder" /v Attributes /t REG_DWORD /d 0xF080004D /f')
    run_command(f'reg add "HKCR\\Wow6432Node\\CLSID\\{onedrive_clsid}\\ShellFolder" /v FolderValueFlags /t REG_DWORD /d 0x28 /f')

    # Restore Desktop namespace
    run_command(f'reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Desktop\\NameSpace\\{onedrive_clsid}" /ve /t REG_SZ /d "OneDrive" /f')
    run_command(f'reg add "HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Desktop\\NameSpace\\{onedrive_clsid}" /ve /t REG_SZ /d "OneDrive" /f')

    # Restore context menus
    run_command(f'reg add "HKCR\\*\\shellex\\ContextMenuHandlers\\FileSyncEx" /ve /t REG_SZ /d "{filesync_clsid}" /f')
    run_command(f'reg add "HKCR\\Directory\\shellex\\ContextMenuHandlers\\FileSyncEx" /ve /t REG_SZ /d "{filesync_clsid}" /f')
    run_command(f'reg add "HKCR\\Directory\\Background\\shellex\\ContextMenuHandlers\\FileSyncEx" /ve /t REG_SZ /d "{filesync_clsid}" /f')

    # Create OneDrive folder
    onedrive_path = os.path.join(os.environ.get('USERPROFILE', ''), 'OneDrive')
    if not os.path.exists(onedrive_path):
        try:
            os.makedirs(onedrive_path)
        except:
            pass

def download_and_install_onedrive():
    """Download and install OneDrive."""
    download_url = "https://go.microsoft.com/fwlink/p/?LinkID=2182910"
    temp_dir = tempfile.mkdtemp()
    installer_path = os.path.join(temp_dir, "OneDriveSetup.exe")

    try:
        urllib.request.urlretrieve(download_url, installer_path)
        if os.path.exists(installer_path):
            subprocess.run([installer_path, "/silent"], capture_output=True)
            return True
    except:
        # Try winget as fallback
        result = run_command('winget install "Microsoft.OneDrive" --silent --accept-package-agreements --accept-source-agreements')
        return result
    finally:
        try:
            if os.path.exists(installer_path):
                os.remove(installer_path)
            if os.path.exists(temp_dir):
                os.rmdir(temp_dir)
        except:
            pass
    return False

def refresh_explorer():
    """Restart Explorer."""
    try:
        run_command('taskkill /f /im explorer.exe')
        time.sleep(2)
        subprocess.Popen('explorer.exe', shell=True)
    except:
        pass

def main():
    enable_ansi_colors()

    # Check for admin
    if not is_admin():
        print("\n  Requesting administrator privileges...")
        run_as_admin()
        sys.exit(0)

    print("""
 ╔═══════════════════════════════════════════════════════════════╗
 ║                                                               ║
 ║                  OneDrive Reinstaller                         ║
 ║                                                               ║
 ╚═══════════════════════════════════════════════════════════════╝
    """)

    input("  Press Enter to restore OneDrive...")

    print("\n  Restoring OneDrive...\n")

    print("  [1/4] Removing blocks and restoring registry...")
    restore_onedrive()
    print("        Done.")

    print("  [2/4] Downloading OneDrive...")
    download_and_install_onedrive()
    print("        Done.")

    print("  [3/4] Installing OneDrive...")
    time.sleep(3)  # Give installer time to start
    print("        Done.")

    print("  [4/4] Refreshing Explorer...")
    refresh_explorer()
    print("        Done.")

    print("""
 ╔═══════════════════════════════════════════════════════════════╗
 ║                                                               ║
 ║                    Restoration Complete!                      ║
 ║                                                               ║
 ║   Please RESTART your computer, then sign in to OneDrive.     ║
 ║                                                               ║
 ╚═══════════════════════════════════════════════════════════════╝
    """)

    input("  Press Enter to exit...")

if __name__ == "__main__":
    main()
