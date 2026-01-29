"""
OneDrive Reinstaller

This tool restores OneDrive functionality after it has been
completely removed by system optimization tools.
"""

import os
import sys
import ctypes
import subprocess
import winreg
import shutil
import urllib.request
import tempfile
import time

# Windows console colors
class Colors:
    HEADER = '\033[95m'
    BLUE = '\033[94m'
    CYAN = '\033[96m'
    GREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    END = '\033[0m'
    BOLD = '\033[1m'

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

def print_banner():
    """Print the application banner."""
    banner = """
 ╔═══════════════════════════════════════════════════════════════╗
 ║                                                               ║
 ║              OneDrive Reinstaller v1.0                        ║
 ║              ─────────────────────────                        ║
 ║                                                               ║
 ║   Restores OneDrive after removal by system tools             ║
 ║                                                               ║
 ╚═══════════════════════════════════════════════════════════════╝
    """
    print(f"{Colors.CYAN}{banner}{Colors.END}")

def print_step(step_num, total, message):
    """Print a step header."""
    print(f"\n{Colors.CYAN}[Step {step_num}/{total}]{Colors.END} {message}")

def print_success(message):
    """Print a success message."""
    print(f"   {Colors.GREEN}[OK]{Colors.END} {message}")

def print_warning(message):
    """Print a warning message."""
    print(f"   {Colors.WARNING}[WARNING]{Colors.END} {message}")

def print_error(message):
    """Print an error message."""
    print(f"   {Colors.FAIL}[ERROR]{Colors.END} {message}")

def print_info(message):
    """Print an info message."""
    print(f"   {Colors.BLUE}[INFO]{Colors.END} {message}")

def run_command(command, silent=True):
    """Run a command and return success status."""
    try:
        if silent:
            result = subprocess.run(
                command,
                shell=True,
                capture_output=True,
                text=True
            )
        else:
            result = subprocess.run(command, shell=True)
        return result.returncode == 0
    except Exception as e:
        return False

def set_registry_value(key_path, value_name, value_data, value_type=winreg.REG_DWORD, hkey=winreg.HKEY_LOCAL_MACHINE):
    """Set a registry value, creating the key if necessary."""
    try:
        # Create key if it doesn't exist
        key = winreg.CreateKeyEx(hkey, key_path, 0, winreg.KEY_ALL_ACCESS)
        winreg.SetValueEx(key, value_name, 0, value_type, value_data)
        winreg.CloseKey(key)
        return True
    except Exception as e:
        return False

def delete_registry_value(key_path, value_name, hkey=winreg.HKEY_LOCAL_MACHINE):
    """Delete a registry value."""
    try:
        key = winreg.OpenKey(hkey, key_path, 0, winreg.KEY_ALL_ACCESS)
        winreg.DeleteValue(key, value_name)
        winreg.CloseKey(key)
        return True
    except FileNotFoundError:
        return True  # Already doesn't exist
    except Exception as e:
        return False

def remove_group_policy_blocks():
    """Remove OneDrive Group Policy blocks."""
    print_step(1, 7, "Removing OneDrive Group Policy blocks...")

    policy_path = r"SOFTWARE\Policies\Microsoft\Windows\OneDrive"

    # Remove blocking values
    delete_registry_value(policy_path, "DisableFileSyncNGSC")
    delete_registry_value(policy_path, "DisableFileSync")
    delete_registry_value(policy_path, "DisableLibrariesDefaultSaveToOneDrive")

    # Set to enabled (0) as backup
    set_registry_value(policy_path, "DisableFileSyncNGSC", 0)
    set_registry_value(policy_path, "DisableFileSync", 0)

    print_success("Group Policy blocks removed")

def restore_explorer_integration():
    """Restore OneDrive Explorer shell integration."""
    print_step(2, 7, "Restoring OneDrive Explorer shell integration...")

    onedrive_clsid = "{018D5C66-4533-4307-9B53-224DE2ED1FE6}"

    # Registry commands for CLSID restoration
    commands = [
        # Main CLSID
        f'reg add "HKCR\\CLSID\\{onedrive_clsid}" /ve /t REG_SZ /d "OneDrive" /f',
        f'reg add "HKCR\\CLSID\\{onedrive_clsid}" /v System.IsPinnedToNameSpaceTree /t REG_DWORD /d 1 /f',
        f'reg add "HKCR\\CLSID\\{onedrive_clsid}\\ShellFolder" /v Attributes /t REG_DWORD /d 0xF080004D /f',
        f'reg add "HKCR\\CLSID\\{onedrive_clsid}\\ShellFolder" /v FolderValueFlags /t REG_DWORD /d 0x28 /f',

        # 64-bit version
        f'reg add "HKCR\\Wow6432Node\\CLSID\\{onedrive_clsid}" /ve /t REG_SZ /d "OneDrive" /f',
        f'reg add "HKCR\\Wow6432Node\\CLSID\\{onedrive_clsid}" /v System.IsPinnedToNameSpaceTree /t REG_DWORD /d 1 /f',
        f'reg add "HKCR\\Wow6432Node\\CLSID\\{onedrive_clsid}\\ShellFolder" /v Attributes /t REG_DWORD /d 0xF080004D /f',
        f'reg add "HKCR\\Wow6432Node\\CLSID\\{onedrive_clsid}\\ShellFolder" /v FolderValueFlags /t REG_DWORD /d 0x28 /f',
    ]

    success_count = sum(1 for cmd in commands if run_command(cmd))

    if success_count == len(commands):
        print_success("Explorer shell integration restored")
    else:
        print_warning(f"Restored {success_count}/{len(commands)} registry entries")

def restore_desktop_namespace():
    """Restore OneDrive in Desktop namespace (File Explorer sidebar)."""
    print_step(3, 7, "Restoring Desktop namespace entries...")

    onedrive_clsid = "{018D5C66-4533-4307-9B53-224DE2ED1FE6}"

    commands = [
        f'reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Desktop\\NameSpace\\{onedrive_clsid}" /ve /t REG_SZ /d "OneDrive" /f',
        f'reg add "HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Desktop\\NameSpace\\{onedrive_clsid}" /ve /t REG_SZ /d "OneDrive" /f',
    ]

    for cmd in commands:
        run_command(cmd)

    print_success("Desktop namespace restored")

def restore_context_menus():
    """Restore OneDrive right-click context menu handlers."""
    print_step(4, 7, "Restoring context menu handlers...")

    filesync_clsid = "{CB3D0F55-BC2C-4C1A-85ED-23ED75B5106B}"

    commands = [
        f'reg add "HKCR\\*\\shellex\\ContextMenuHandlers\\FileSyncEx" /ve /t REG_SZ /d "{filesync_clsid}" /f',
        f'reg add "HKCR\\Directory\\shellex\\ContextMenuHandlers\\FileSyncEx" /ve /t REG_SZ /d "{filesync_clsid}" /f',
        f'reg add "HKCR\\Directory\\Background\\shellex\\ContextMenuHandlers\\FileSyncEx" /ve /t REG_SZ /d "{filesync_clsid}" /f',
    ]

    for cmd in commands:
        run_command(cmd)

    print_success("Context menu handlers restored")

def create_onedrive_folder():
    """Create the OneDrive user folder if it doesn't exist."""
    print_step(5, 7, "Creating OneDrive user folder...")

    onedrive_path = os.path.join(os.environ.get('USERPROFILE', ''), 'OneDrive')

    if not os.path.exists(onedrive_path):
        try:
            os.makedirs(onedrive_path)
            print_success(f"Created OneDrive folder at {onedrive_path}")
        except Exception as e:
            print_warning(f"Could not create folder: {e}")
    else:
        print_success("OneDrive folder already exists")

def download_and_install_onedrive():
    """Download and install OneDrive from Microsoft."""
    print_step(6, 7, "Downloading and installing OneDrive...")

    download_url = "https://go.microsoft.com/fwlink/p/?LinkID=2182910"
    temp_dir = tempfile.mkdtemp()
    installer_path = os.path.join(temp_dir, "OneDriveSetup.exe")

    print_info("Downloading OneDrive installer...")

    try:
        # Download with progress
        urllib.request.urlretrieve(download_url, installer_path)

        if os.path.exists(installer_path):
            print_success("Download complete")
            print_info("Installing OneDrive (this may take a moment)...")

            # Run installer silently
            result = subprocess.run(
                [installer_path, "/silent"],
                capture_output=True
            )

            print_success("OneDrive installation initiated")
        else:
            raise Exception("Download failed")

    except Exception as e:
        print_warning(f"Direct download failed: {e}")
        print_info("Trying winget...")

        # Try winget as fallback
        result = run_command(
            'winget install "Microsoft.OneDrive" --silent --accept-package-agreements --accept-source-agreements'
        )

        if result:
            print_success("Installed via winget")
        else:
            print_error("Automatic installation failed")
            print(f"""
   {Colors.WARNING}Please install OneDrive manually:{Colors.END}
   - Download: https://www.microsoft.com/en-us/microsoft-365/onedrive/download
   - Or search "OneDrive" in Microsoft Store
            """)

    finally:
        # Cleanup
        try:
            if os.path.exists(installer_path):
                os.remove(installer_path)
            if os.path.exists(temp_dir):
                os.rmdir(temp_dir)
        except:
            pass

def refresh_explorer():
    """Restart Explorer to apply changes."""
    print_step(7, 7, "Refreshing Explorer...")

    try:
        # Kill explorer
        run_command('taskkill /f /im explorer.exe')
        time.sleep(2)
        # Restart explorer
        subprocess.Popen('explorer.exe', shell=True)
        print_success("Explorer refreshed")
    except Exception as e:
        print_warning(f"Could not restart Explorer: {e}")

def print_completion():
    """Print completion message with next steps."""
    message = f"""
{Colors.GREEN}
 ╔═══════════════════════════════════════════════════════════════╗
 ║                                                               ║
 ║         OneDrive Reinstallation Complete!                     ║
 ║                                                               ║
 ╚═══════════════════════════════════════════════════════════════╝
{Colors.END}
 What to do next:
   1. OneDrive should now be installing in the background
   2. Look for the OneDrive cloud icon in your system tray
   3. Click it to sign in with your Microsoft account
   4. If you don't see it, {Colors.CYAN}restart your computer{Colors.END}

 If OneDrive still won't install:
   - Download manually: {Colors.BLUE}https://onedrive.live.com/about/download{Colors.END}
   - Or search "OneDrive" in Microsoft Store

    """
    print(message)

def main():
    """Main entry point."""
    # Enable colors
    enable_ansi_colors()

    # Check for admin privileges
    if not is_admin():
        print(f"\n{Colors.FAIL}ERROR: Administrator privileges required!{Colors.END}")
        print("Restarting with admin privileges...")
        run_as_admin()
        sys.exit(0)

    # Print banner
    print_banner()

    print(" This tool will:")
    print("   1. Remove OneDrive blocking policies")
    print("   2. Restore Explorer integration")
    print("   3. Restore context menus")
    print("   4. Download and install OneDrive")
    print()

    input(" Press Enter to continue...")

    # Execute restoration steps
    remove_group_policy_blocks()
    restore_explorer_integration()
    restore_desktop_namespace()
    restore_context_menus()
    create_onedrive_folder()
    download_and_install_onedrive()
    refresh_explorer()

    # Print completion
    print_completion()

    input(" Press Enter to exit...")

if __name__ == "__main__":
    main()
