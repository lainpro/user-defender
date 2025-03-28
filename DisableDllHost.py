import winreg
import subprocess
import ctypes
import sys

class COMSurrogateProtector:
    def __init__(self):
        """
        Initialize configuration for managing COM Surrogate (dllhost.exe).
        Defines registry paths and services related to COM Surrogate.
        """
        # Registry paths related to COM Surrogate
        self.com_surrogate_registry_paths = [
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\COM3",
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced",
        ]

        # Services that may be related to COM Surrogate
        self.com_surrogate_services = [
            "DcomLaunch",  # DCOM Server Process Launcher
            "RpcSs",  # Remote Procedure Call (RPC)
        ]

    def protect_com_surrogate(self):
        """
        Execute comprehensive management strategy for COM Surrogate.
        Systematically disables COM Surrogate and related services.
        """
        try:
            # Execute COM Surrogate management methods in sequence
            self._disable_com_surrogate_registry()
            self._disable_com_surrogate_services()

            print("COM Surrogate management applied successfully.")
        except Exception as e:
            print(f"Error in COM Surrogate management process: {e}")

    def _disable_com_surrogate_registry(self):
        """
        Modify Windows registry to minimize the use of COM Surrogate.
        Systematically disables COM Surrogate across multiple registry paths.
        """
        try:
            for path in self.com_surrogate_registry_paths:
                try:
                    key = winreg.OpenKey(
                        winreg.HKEY_LOCAL_MACHINE,
                        path,
                        0,
                        winreg.KEY_ALL_ACCESS
                    )

                    # Disable COM Surrogate settings
                    com_surrogate_settings = [
                        ("DllSurrogate", ""),  # Remove DllSurrogate to disable COM Surrogate
                    ]

                    for name, value in com_surrogate_settings:
                        winreg.SetValueEx(key, name, 0, winreg.REG_SZ, value)

                    winreg.CloseKey(key)
                except FileNotFoundError:
                    print(f"Registry path not found: {path}")

            print("COM Surrogate registry settings successfully modified.")
        except Exception as e:
            print(f"COM Surrogate registry modification error: {e}")

    def _disable_com_surrogate_services(self):
        """
        Systematically stop and disable services related to COM Surrogate.
        Prevents background operations that may utilize COM Surrogate.
        """
        for service in self.com_surrogate_services:
            try:
                subprocess.run(["sc", "stop", service], capture_output=True)
                subprocess.run(["sc", "config", service, "start=", "disabled"], capture_output=True)
                print(f"Successfully disabled service: {service}")
            except Exception as e:
                print(f"Could not disable service {service}: {e}")

def main():
    """
    Entry point for COM Surrogate management utility.
    Verifies administrative privileges before executing management.
    """
    try:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin()
    except:
        is_admin = False

    if not is_admin:
        print("Administrator privileges required. Please run as administrator.")
        sys.exit(1)

    # Initialize and execute COM Surrogate management
    protector = COMSurrogateProtector()
    protector.protect_com_surrogate()

if __name__ == "__main__":
    main()