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


def disable_windows_update_service():
    """Disable Windows Update Service"""
    try:
        # Stop the Windows Update service
        subprocess.run(['net', 'stop', 'wuauserv'], capture_output=True)

        # Disable the service via sc command
        subprocess.run(['sc', 'config', 'wuauserv', 'start=', 'disabled'], capture_output=True)

        print("Windows Update Service disabled.")
    except Exception as e:
        print(f"Error disabling service: {e}")


def disable_group_policy_updates():
    """Disable Windows Updates via Group Policy"""
    try:
        # Use PowerShell to modify Group Policy
        subprocess.run([
            'powershell',
            'Set-ItemProperty',
            '-Path',
            '"HKLM:\\SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsUpdate\\AU"',
            '-Name',
            '"NoAutoUpdate"',
            '-Value',
            '1'
        ], capture_output=True)

        print("Group Policy Updates disabled.")
    except Exception as e:
        print(f"Error modifying Group Policy: {e}")


def disable_windows_update_registry():
    """Disable Windows Updates via Registry"""
    try:
        # Open or create the Windows Update policy key
        key_path = r"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
        key = winreg.CreateKey(
            winreg.HKEY_LOCAL_MACHINE,
            key_path
        )

        # Set AUOptions to 2 (disable automatic updates)
        winreg.SetValueEx(
            key,
            "AUOptions",
            0,
            winreg.REG_DWORD,
            2
        )

        winreg.CloseKey(key)
        print("Registry update settings disabled.")

    except Exception as e:
        print(f"Error modifying registry: {e}")


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

    print("Windows Update Disabler")

    # Disable updates using multiple methods
    disable_windows_update_service()
    disable_group_policy_updates()
    disable_windows_update_registry()

    print("All Windows Update methods disabled.")


if __name__ == "__main__":
    main()
