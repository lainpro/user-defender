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


def disable_start_menu_suggestions():
    """Disable Start Menu suggestions via Registry"""
    try:
        # Registry paths for Start Menu suggestions
        key_paths = [
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
        ]

        for path in key_paths:
            try:
                key = winreg.OpenKey(
                    winreg.HKEY_CURRENT_USER,
                    path,
                    0,
                    winreg.KEY_ALL_ACCESS
                )

                # Disable various suggestion settings
                suggestion_keys = [
                    "ContentDeliveryAllowed",
                    "OemPreInstalledAppsEnabled",
                    "PreInstalledAppsEnabled",
                    "PreInstalledAppsEverEnabled",
                    "SilentInstalledAppsEnabled",
                    "SystemPaneSuggestionsEnabled",
                    "SubscribedContent-338388Enabled",
                    "SubscribedContent-338389Enabled",
                    "SubscribedContent-338393Enabled",
                    "SubscribedContent-353694Enabled",
                    "SubscribedContent-353696Enabled"
                ]

                for suggestion_key in suggestion_keys:
                    winreg.SetValueEx(
                        key,
                        suggestion_key,
                        0,
                        winreg.REG_DWORD,
                        0
                    )

                winreg.CloseKey(key)
            except Exception as e:
                print(f"Error modifying {path}: {e}")

        print("Start Menu suggestions disabled.")
    except Exception as e:
        print(f"Error disabling Start Menu suggestions: {e}")


def disable_lock_screen_ads():
    """Disable Lock Screen ads and suggestions"""
    try:
        key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            key_path,
            0,
            winreg.KEY_ALL_ACCESS
        )

        # Disable lock screen suggestions
        lock_screen_keys = [
            "RotatingLockScreenEnabled",
            "RotatingLockScreenOverlayEnabled",
            "SubscribedContent-338387Enabled"
        ]

        for lock_key in lock_screen_keys:
            winreg.SetValueEx(
                key,
                lock_key,
                0,
                winreg.REG_DWORD,
                0
            )

        winreg.CloseKey(key)
        print("Lock Screen ads disabled.")
    except Exception as e:
        print(f"Error disabling Lock Screen ads: {e}")


def disable_file_explorer_ads():
    """Disable File Explorer and OneDrive ads"""
    try:
        # Disable OneDrive startup
        startup_key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
        startup_key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            startup_key_path,
            0,
            winreg.KEY_ALL_ACCESS
        )

        # Remove OneDrive from startup
        try:
            winreg.DeleteValue(startup_key, "OneDrive")
        except FileNotFoundError:
            pass

        winreg.CloseKey(startup_key)

        # Disable OneDrive notifications via PowerShell
        subprocess.run([
            'powershell',
            'New-ItemProperty',
            '-Path',
            'HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\CloudStore\\Store\\DefaultAccount',
            '-Name',
            '"AutobackupNotSupportedNotificationFlag"',
            '-Value',
            '1',
            '-PropertyType',
            'DWord'
        ], capture_output=True)

        print("File Explorer and OneDrive ads disabled.")
    except Exception as e:
        print(f"Error disabling File Explorer ads: {e}")


def disable_cortana_suggestions():
    """Disable Cortana and Search bar suggestions"""
    try:
        # Disable Cortana suggestions in registry
        key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Search"
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            key_path,
            0,
            winreg.KEY_ALL_ACCESS
        )

        # Disable various search and Cortana suggestions
        suggestion_keys = [
            "SearchboxTaskbarMode",  # Hide search bar
            "BingSearchEnabled",  # Disable Bing searches
            "CortanaConsent"  # Disable Cortana
        ]

        for suggestion_key in suggestion_keys:
            winreg.SetValueEx(
                key,
                suggestion_key,
                0,
                winreg.REG_DWORD,
                0
            )

        winreg.CloseKey(key)
        print("Cortana and Search suggestions disabled.")
    except Exception as e:
        print(f"Error disabling Cortana suggestions: {e}")


def disable_notifications_ads():
    """Disable notification ads"""
    try:
        # Modify Windows Notification settings
        key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Notifications\Settings"
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            key_path,
            0,
            winreg.KEY_ALL_ACCESS
        )

        # Disable notifications for specific system and advertising apps
        notification_keys = [
            "Microsoft.Windows.CloudExperienceHost_cw5n1h2txyewy",
            "Microsoft.Windows.ShellExperienceHost_cw5n1h2txyewy"
        ]

        for app in notification_keys:
            try:
                winreg.SetValueEx(
                    key,
                    app,
                    0,
                    winreg.REG_DWORD,
                    0
                )
            except Exception:
                pass

        winreg.CloseKey(key)
        print("Notification ads disabled.")
    except Exception as e:
        print(f"Error disabling notification ads: {e}")


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

    print("Windows Ads and Suggestions Disabler")

    # Disable ads and suggestions using multiple methods
    disable_start_menu_suggestions()
    disable_lock_screen_ads()
    disable_file_explorer_ads()
    disable_cortana_suggestions()
    disable_notifications_ads()

    print("All ads and suggestions have been disabled.")
    print("Restart your computer to apply all changes.")


if __name__ == "__main__":
    main()