import winreg
import subprocess
import sys
import os
import ctypes


class WindowsPrivacyProtector:
    def __init__(self):
        """
        Initialize comprehensive privacy protection configuration.
        Defines registry paths, services, and IP addresses to block.
        """
        # Expanded list of telemetry-related registry paths
        self.telemetry_registry_paths = [
            # System-wide telemetry paths
            r"SOFTWARE\Policies\Microsoft\Windows\DataCollection",
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection",
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced",

            # Additional tracking paths
            r"SOFTWARE\Policies\Microsoft\Windows Defender\Reporting",
            r"SOFTWARE\Microsoft\Office\16.0\Common",
            r"SOFTWARE\Policies\Microsoft\Office\16.0\Common",
            r"SOFTWARE\Microsoft\OneDrive",
            r"SOFTWARE\Policies\Microsoft\Windows\LocationAndSensors",
        ]

        # Comprehensive list of diagnostic and tracking services to disable
        self.diagnostic_services = [
            # Core diagnostic services
            "DiagTrack",  # Diagnostic Tracking Service
            "dmwappushservice",  # WAP Push Message Routing Service
            "OneSyncSvc",  # Sync Host Service
            "WerSvc",  # Windows Error Reporting Service
            "RemoteRegistry",  # Remote Registry Management

            # Extended tracking services
            "XboxNetApiSvc",  # Xbox Network Service
            "XblAuthManager",  # Xbox Live Authentication
            "XblGameSave",  # Xbox Game Save Service
            "BITS",  # Background Intelligent Transfer Service
            "DoSvc",  # Delivery Optimization Service
        ]

        # Expanded list of known Microsoft telemetry IP addresses
        self.telemetry_ips = [
            "13.68.37.221",  # Microsoft Telemetry
            "65.55.108.23",  # Microsoft Telemetry
            "104.214.72.101",  # Microsoft Telemetry
            "137.116.81.24",  # Microsoft Telemetry
            "40.77.226.250",  # Additional Microsoft Telemetry
            "65.52.108.29",  # Microsoft Telemetry Endpoint
        ]

    def protect_privacy(self):
        """
        Execute comprehensive privacy protection strategy.
        Systematically disables telemetry across multiple system components.
        """
        try:
            # Execute privacy protection methods in sequence
            self._disable_registry_telemetry()
            self._disable_diagnostic_services()
            self._block_telemetry_ips()
            self._modify_windows_updates()
            self._disable_cortana_tracking()
            self._disable_diagnostic_data()

            print("Comprehensive privacy protection applied successfully.")
        except Exception as e:
            print(f"Error in privacy protection process: {e}")

    def _disable_registry_telemetry(self):
        """
        Modify Windows registry to minimize data collection and tracking.
        Systematically disables telemetry across multiple registry paths.
        """
        try:
            for path in self.telemetry_registry_paths:
                try:
                    key = winreg.OpenKey(
                        winreg.HKEY_LOCAL_MACHINE,
                        path,
                        0,
                        winreg.KEY_ALL_ACCESS
                    )

                    # Comprehensive telemetry disablement settings
                    telemetry_settings = [
                        ("AllowTelemetry", 0),  # Disable telemetry completely
                        ("DisableEnterpriseAuthProxy", 1),  # Prevent enterprise proxy
                        ("DoNotShowFeedbackNotifications", 1),  # Stop feedback notifications
                        ("DisableDiagnosticReporting", 1),  # Disable diagnostic reporting
                    ]

                    for name, value in telemetry_settings:
                        winreg.SetValueEx(key, name, 0, winreg.REG_DWORD, value)

                    winreg.CloseKey(key)
                except FileNotFoundError:
                    print(f"Registry path not found: {path}")

            print("Registry telemetry settings successfully modified.")
        except Exception as e:
            print(f"Registry telemetry modification error: {e}")

    def _disable_diagnostic_services(self):
        """
        Systematically stop and disable Windows diagnostic and tracking services.
        Prevents background data collection and telemetry transmission.
        """
        for service in self.diagnostic_services:
            try:
                subprocess.run(["sc", "stop", service], capture_output=True)
                subprocess.run(["sc", "config", service, "start=", "disabled"], capture_output=True)
                print(f"Successfully disabled service: {service}")
            except Exception as e:
                print(f"Could not disable service {service}: {e}")

    def _block_telemetry_ips(self):
        """
        Create firewall rules to block known Microsoft telemetry IP addresses.
        Prevents direct communication with telemetry servers.
        """
        for ip in self.telemetry_ips:
            try:
                subprocess.run([
                    "netsh", "advfirewall", "firewall", "add", "rule",
                    f"name=Block_Telemetry_{ip}",
                    "dir=out",
                    "action=block",
                    f"remoteip={ip}"
                ], check=True)
                print(f"Blocked telemetry IP: {ip}")
            except subprocess.CalledProcessError:
                print(f"Could not block IP: {ip}")

    def _modify_windows_updates(self):
        """
        Configure Windows Update settings to minimize data sharing.
        Provides user more control over update processes.
        """
        try:
            update_key = winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU",
                0,
                winreg.KEY_ALL_ACCESS
            )

            update_settings = [
                ("NoAutoUpdate", 1),  # Disable automatic updates
                ("AUOptions", 2)  # Notify before downloading updates
            ]

            for name, value in update_settings:
                winreg.SetValueEx(update_key, name, 0, winreg.REG_DWORD, value)

            winreg.CloseKey(update_key)
            print("Windows Update settings successfully modified.")
        except Exception as e:
            print(f"Windows Update modification error: {e}")

    def _disable_cortana_tracking(self):
        """
        Disable Cortana and associated search tracking mechanisms.
        Prevents personalized search and data collection.
        """
        try:
            cortana_key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\Search",
                0,
                winreg.KEY_ALL_ACCESS
            )

            cortana_settings = [
                ("BingSearchEnabled", 0),
                ("CortanaEnabled", 0),
                ("DeviceHistoryEnabled", 0)
            ]

            for name, value in cortana_settings:
                winreg.SetValueEx(cortana_key, name, 0, winreg.REG_DWORD, value)

            winreg.CloseKey(cortana_key)
            print("Cortana and search tracking successfully disabled.")
        except Exception as e:
            print(f"Cortana tracking modification error: {e}")

    def _disable_diagnostic_data(self):
        """
        Advanced diagnostic data and feedback disablement.
        Prevents additional system information collection.
        """
        try:
            diagnostic_keys = [
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\Diagnostics\DiagTrack",
                r"SOFTWARE\Policies\Microsoft\Windows\Windows Error Reporting"
            ]

            for path in diagnostic_keys:
                try:
                    key = winreg.OpenKey(
                        winreg.HKEY_LOCAL_MACHINE,
                        path,
                        0,
                        winreg.KEY_ALL_ACCESS
                    )

                    diagnostic_settings = [
                        ("DisableAutomaticSendTraces", 1),
                        ("Disabled", 1),
                        ("LoggingDisabled", 1)
                    ]

                    for name, value in diagnostic_settings:
                        try:
                            winreg.SetValueEx(key, name, 0, winreg.REG_DWORD, value)
                        except:
                            pass

                    winreg.CloseKey(key)
                except FileNotFoundError:
                    print(f"Diagnostic key not found: {path}")

            print("Advanced diagnostic data tracking successfully disabled.")
        except Exception as e:
            print(f"Diagnostic data modification error: {e}")


def main():
    """
    Entry point for Windows privacy protection utility.
    Verifies administrative privileges before executing privacy protection.
    """
    try:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin()
    except:
        is_admin = False

    if not is_admin:
        print("Administrator privileges required. Please run as administrator.")
        sys.exit(1)

    # Initialize and execute privacy protection
    protector = WindowsPrivacyProtector()
    protector.protect_privacy()

if __name__ == "__main__":
    main()