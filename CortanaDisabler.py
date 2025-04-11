import winreg
import os
import sys
import ctypes
import subprocess


def is_admin():
    """Check if script is running with administrative privileges."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except:
        return False


def disable_cortana_registry():
    """Disable Cortana through Windows Registry"""
    try:
        # Path to disable Cortana in Windows Registry
        key_paths = [
            r"SOFTWARE\Policies\Microsoft\Windows\Windows Search",
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Search"
        ]

        for path in key_paths:
            try:
                # Open or create the key
                key = winreg.CreateKey(
                    winreg.HKEY_LOCAL_MACHINE,
                    path
                )

                # Disable Cortana and Windows Search
                winreg.SetValueEx(
                    key,
                    "AllowCortana",
                    0,
                    winreg.REG_DWORD,
                    0
                )

                winreg.SetValueEx(
                    key,
                    "DisableSearch",
                    0,
                    winreg.REG_DWORD,
                    1
                )

                winreg.CloseKey(key)
            except Exception as e:
                print(f"Error modifying {path}: {e}")

        print("Cortana Registry settings disabled.")
    except Exception as e:
        print(f"Error in registry modification: {e}")


def disable_cortana_services():
    """Disable Cortana-related Windows Services"""
    cortana_services = [
        "WSearch",  # Windows Search service
        "WSearchIndexer"  # Windows Search Indexer
    ]

    for service in cortana_services:
        try:
            # Stop the service
            subprocess.run(['net', 'stop', service], capture_output=True)

            # Disable the service
            subprocess.run(['sc', 'config', service, 'start=', 'disabled'], capture_output=True)
        except Exception as e:
            print(f"Error disabling {service}: {e}")

    print("Cortana-related services disabled.")


def remove_cortana_taskbar():
    """Remove Cortana from Taskbar via Registry"""
    try:
        key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            key_path,
            0,
            winreg.KEY_ALL_ACCESS
        )

        # Set value to hide Cortana from taskbar
        winreg.SetValueEx(
            key,
            "ShowCortanaButton",
            0,
            winreg.REG_DWORD,
            0
        )

        winreg.CloseKey(key)
        print("Cortana taskbar icon removed.")
    except Exception as e:
        print(f"Error removing Cortana from taskbar: {e}")


def main():
    # Check for admin rights
    if not is_admin():
        # Re-run the script with admin rights
        ctypes.windll.shell32.ShellExecuteW(
            None,
            "runas",
            sys.executable,
            " ".join(sys.argv),
            None,
            1
        )
        sys.exit()

    print("Cortana Disabler Script")

    # Disable Cortana using multiple methods
    disable_cortana_registry()
    disable_cortana_services()
    remove_cortana_taskbar()

    print("Cortana has been disabled.")
    print("Restart your computer to apply all changes.")


if __name__ == "__main__":
    main()