#!/usr/bin/env python3
"""
Windows Image Preparation GUI
A comprehensive tool for preparing Windows images for generalization
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import subprocess
import threading
import re
import os
import platform
import ctypes
import requests
from pathlib import Path
import string
import time

def check_platform():
    """Check if running on Windows"""
    if platform.system() != 'Windows':
        print(f"Error: This tool is designed for Windows systems only.")
        print(f"Current platform: {platform.system()}")
        print("Please run this tool on a Windows machine where you're preparing the image.")
        return False
    return True

class WindowsImagePrepGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("OS Imaging Tool")
        self.root.geometry("700x600")
        self.root.minsize(600, 500)

        # Style
        self.style = ttk.Style(self.root)
        self.style.theme_use('vista')
        
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        # --- Configuration Section ---
        config_frame = ttk.LabelFrame(main_frame, text="Configuration", padding="10")
        config_frame.pack(fill="x", expand=False)
        config_frame.columnconfigure(1, weight=1)

        # Disk2vhd Path
        ttk.Label(config_frame, text="Disk2vhd Path:").grid(row=0, column=0, sticky="w", pady=2)
        self.disk2vhd_path_var = tk.StringVar(value=str(Path("./disk2vhd.exe").resolve()))
        self.disk2vhd_entry = ttk.Entry(config_frame, textvariable=self.disk2vhd_path_var, state='readonly')
        self.disk2vhd_entry.grid(row=0, column=1, sticky="we", padx=5)
        self.download_button = ttk.Button(config_frame, text="Check / Download", command=self.check_and_download_disk2vhd)
        self.download_button.grid(row=0, column=2, sticky="e")

        # Destination Path
        ttk.Label(config_frame, text="VHDX UNC Path:").grid(row=1, column=0, sticky="w", pady=2)
        self.unc_path_var = tk.StringVar()
        self.unc_entry = ttk.Entry(config_frame, textvariable=self.unc_path_var)
        self.unc_entry.grid(row=1, column=1, columnspan=2, sticky="we", padx=5)
        self.unc_entry.insert(0, "\\\\server\\share\\image-name.vhdx")
        
        # Credentials
        ttk.Label(config_frame, text="Username:").grid(row=2, column=0, sticky="w", pady=2)
        self.username_var = tk.StringVar()
        self.username_entry = ttk.Entry(config_frame, textvariable=self.username_var)
        self.username_entry.grid(row=2, column=1, columnspan=2, sticky="we", padx=5)
        
        ttk.Label(config_frame, text="Password:").grid(row=3, column=0, sticky="w", pady=2)
        self.password_var = tk.StringVar()
        self.password_entry = ttk.Entry(config_frame, textvariable=self.password_var, show="*")
        self.password_entry.grid(row=3, column=1, columnspan=2, sticky="we", padx=5)

        # --- Action Button ---
        self.create_button = ttk.Button(main_frame, text="Create VHDX Image", command=self.start_image_creation_thread)
        self.create_button.pack(pady=10, fill="x")
        
        # --- Generalization Section ---
        generalize_frame = ttk.LabelFrame(main_frame, text="Generalization (Sysprep)", padding="10")
        generalize_frame.pack(fill="x", expand=False, pady=5)
        generalize_frame.columnconfigure(0, weight=1)

        self.skip_user_cleanup_var = tk.BooleanVar()
        self.skip_agent_cleanup_var = tk.BooleanVar()
        self.skip_log_cleanup_var = tk.BooleanVar()

        ttk.Checkbutton(generalize_frame, text="Skip User Profile & Account Cleanup", variable=self.skip_user_cleanup_var).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(generalize_frame, text="Skip Agent Identity Cleanup (e.g., NinjaRMM)", variable=self.skip_agent_cleanup_var).grid(row=1, column=0, sticky="w")
        ttk.Checkbutton(generalize_frame, text="Skip System Log Cleanup (Event Logs, etc.)", variable=self.skip_log_cleanup_var).grid(row=2, column=0, sticky="w")
        
        self.generalize_button = ttk.Button(generalize_frame, text="Prepare and Generalize System", command=self.start_generalization_thread)
        self.generalize_button.grid(row=3, column=0, sticky="ew", pady=10)

        # --- Log Area ---
        log_frame = ttk.LabelFrame(main_frame, text="Log", padding="10")
        log_frame.pack(fill="both", expand=True)
        self.log_area = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, state='disabled', bg="#f0f0f0")
        self.log_area.pack(fill="both", expand=True)
        
        # --- Initial Checks ---
        self.check_admin()
        self.check_and_download_disk2vhd()

    def log(self, message):
        """Appends a message to the log area in a thread-safe way."""
        def _append():
            self.log_area.config(state='normal')
            self.log_area.insert(tk.END, message + "\n")
            self.log_area.config(state='disabled')
            self.log_area.see(tk.END)
        self.root.after(0, _append)

    def check_admin(self):
        """Check for admin rights and log the result."""
        try:
            is_admin = ctypes.windll.shell32.IsUserAnAdmin()
        except Exception as e:
            self.log(f"Admin check failed: {e}")
            is_admin = False

        if not is_admin:
            self.log("ERROR: This application must be run as Administrator.")
            messagebox.showerror("Permission Denied", "Please restart the application with Administrator privileges.")
            self.root.destroy()
            return False
        else:
            self.log("INFO: Running with Administrator privileges.")
        return is_admin

    def check_and_download_disk2vhd(self):
        """Checks if disk2vhd.exe exists, and downloads it if not."""
        self.log("INFO: Checking for disk2vhd.exe...")
        disk2vhd_path = Path(self.disk2vhd_path_var.get())
        if disk2vhd_path.exists():
            self.log("SUCCESS: disk2vhd.exe found.")
            return True
        
        self.log("INFO: disk2vhd.exe not found. Attempting to download...")
        url = "https://live.sysinternals.com/disk2vhd.exe"
        try:
            with requests.get(url, stream=True) as r:
                r.raise_for_status()
                with open(disk2vhd_path, 'wb') as f:
                    for chunk in r.iter_content(chunk_size=8192):
                        f.write(chunk)
            self.log(f"SUCCESS: Downloaded disk2vhd.exe to {disk2vhd_path}")
            return True
        except Exception as e:
            self.log(f"ERROR: Failed to download disk2vhd.exe: {e}")
            messagebox.showerror("Download Failed", f"Could not download disk2vhd.exe.\nPlease download it manually from Sysinternals and place it next to this application.")
            return False

    def find_unused_drive_letter(self):
        """Finds an available drive letter, starting from Z:."""
        for letter in reversed(string.ascii_uppercase):
            drive = f"{letter}:\\"
            if not os.path.exists(drive):
                self.log(f"INFO: Found unused drive letter: {letter}:")
                return f"{letter}:"
        return None

    def start_image_creation_thread(self):
        """Starts the imaging process in a new thread to avoid freezing the GUI."""
        self.create_button.config(state="disabled")
        thread = threading.Thread(target=self.create_image_worker)
        thread.daemon = True
        thread.start()

    def create_image_worker(self):
        """The main worker function that performs all imaging tasks."""
        try:
            # 1. Get parameters from GUI
            unc_path = self.unc_path_var.get()
            username = self.username_var.get()
            password = self.password_var.get()
            disk2vhd_exe = self.disk2vhd_path_var.get()

            if not unc_path or not unc_path.startswith("\\\\"):
                self.log("ERROR: Invalid UNC path provided. It must start with \\\\.")
                messagebox.showerror("Invalid Input", "Please provide a valid UNC path (e.g., \\\\server\\share\\image.vhdx).")
                return

            # 2. Find OS volumes
            self.log("INFO: Identifying OS volumes to capture...")
            try:
                ps_command_disk = "(Get-Partition | Where-Object { $_.DriveLetter -eq $env:SystemDrive.Trim(':') }).DiskNumber"
                disk_number_result = subprocess.run(["powershell", "-Command", ps_command_disk], capture_output=True, text=True, check=True)
                disk_number = disk_number_result.stdout.strip()

                ps_command_volumes = f"(Get-Partition -DiskNumber {disk_number} | ForEach-Object {{ $_.DriveLetter }} | Where-Object {{ -not [string]::IsNullOrWhiteSpace($_) }})"
                volumes_result = subprocess.run(["powershell", "-Command", ps_command_volumes], capture_output=True, text=True, check=True)
                volumes_to_capture = [v.strip() + ":" for v in volumes_result.stdout.strip().splitlines() if v.strip()]
                
                if not volumes_to_capture:
                    raise ValueError("Could not auto-detect any volumes.")

                self.log(f"INFO: Volumes to be captured: {', '.join(volumes_to_capture)}")
            except (subprocess.CalledProcessError, ValueError) as e:
                self.log(f"ERROR: Failed to identify OS volumes: {e}")
                return
            
            # 3. Map Network Drive
            drive_letter = self.find_unused_drive_letter()
            if not drive_letter:
                self.log("ERROR: Could not find an unused drive letter to map the network share.")
                return
            
            unc_path_obj = Path(unc_path)
            share_path = str(unc_path_obj.parent)
            image_filename = unc_path_obj.name
            local_path = f"{drive_letter}\\{image_filename}"
            
            self.log(f"INFO: Mapping network share '{share_path}' to drive {drive_letter}...")
            net_use_cmd = ["net", "use", drive_letter, share_path]
            if username and password:
                net_use_cmd.append(f"/user:{username}")
                net_use_cmd.append(password)
            
            map_proc = subprocess.run(net_use_cmd, capture_output=True, text=True)
            if map_proc.returncode != 0:
                self.log(f"ERROR: Failed to map network drive. {map_proc.stderr or map_proc.stdout}")
                return
            self.log("SUCCESS: Network drive mapped successfully.")

            # 4. Run Disk2vhd
            self.log("INFO: Starting Disk2vhd capture. This may take a long time...")
            arguments = ["-accepteula", "-v", "-q", "-p"] + volumes_to_capture + [local_path]
            self.log(f"COMMAND: \"{disk2vhd_exe}\" {' '.join(arguments)}")

            process = subprocess.Popen(
                [disk2vhd_exe] + arguments, 
                stdout=subprocess.PIPE, 
                stderr=subprocess.STDOUT, 
                text=True, 
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            if process.stdout:
                for line in iter(process.stdout.readline, ''):
                    self.log(line.strip())
            
            process.wait()

            if process.returncode == 0:
                self.log("SUCCESS: Disk2vhd completed successfully!")
                self.log(f"SUCCESS: Image saved to: {unc_path}")
            else:
                self.log(f"ERROR: Disk2vhd failed with exit code: {process.returncode}.")

        except Exception as e:
            self.log(f"FATAL: An unexpected error occurred: {e}")
        finally:
            # 5. Cleanup
            self.log("INFO: Cleaning up mapped network drive...")
            drive_letter_to_unmap = locals().get("drive_letter")
            if drive_letter_to_unmap and os.path.exists(drive_letter_to_unmap + "\\"):
                unmap_proc = subprocess.run(["net", "use", drive_letter_to_unmap, "/delete"], capture_output=True, text=True)
                if unmap_proc.returncode == 0:
                    self.log(f"INFO: Network drive {drive_letter_to_unmap} has been removed.")
                else:
                    self.log(f"WARN: Failed to remove network drive {drive_letter_to_unmap}. {unmap_proc.stderr or unmap_proc.stdout}")
            self.create_button.config(state="normal")

    def run_powershell(self, command, title):
        """Runs a PowerShell command and logs the output."""
        self.log(f"--- Running: {title} ---")
        try:
            full_command = f"$ProgressPreference = 'SilentlyContinue'; {command}"
            process = subprocess.Popen(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", full_command],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            if process.stdout:
                for line in iter(process.stdout.readline, ''):
                    if line.strip():
                        self.log(line.strip())
            process.wait()
            if process.returncode == 0:
                self.log(f"--- SUCCESS: {title} completed. ---")
                return True
            else:
                self.log(f"--- ERROR: {title} failed with exit code {process.returncode}. ---")
                return False
        except Exception as e:
            self.log(f"--- FATAL ERROR in {title}: {e} ---")
            return False

    def start_generalization_thread(self):
        """Starts the generalization process in a new thread."""
        if not messagebox.askyesno("Confirm Generalization", 
                                   "This will prepare the system for cloning using Sysprep. The process is irreversible and will SHUT DOWN the computer upon completion. Are you sure you want to continue?"):
            return
        
        self.generalize_button.config(state="disabled")
        self.create_button.config(state="disabled")
        thread = threading.Thread(target=self.generalize_worker)
        thread.daemon = True
        thread.start()

    def generalize_worker(self):
        """Worker function that performs all generalization tasks, mirroring the PowerShell script."""
        try:
            # Replicating the main execution block from the PowerShell script
            self.log("INFO: Starting Windows Image Preparation...")

            # STEP 1: Remove User Accounts and Profiles
            if not self.skip_user_cleanup_var.get():
                self.run_powershell(
                    title="STEP 1: Remove User Accounts and Profiles",
                    command="""
                        $currentUser = ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\\')[-1]
                        $systemAccounts = @('Administrator','DefaultAccount','Guest','WDAGUtilityAccount', $currentUser)
                        $systemProfiles = @('Administrator','DefaultAccount','Guest','WDAGUtilityAccount','Public', $currentUser)
                        
                        Write-Host "--- Removing local user accounts ---"
                        Get-LocalUser | Where-Object { $_.Name -notin $systemAccounts } | ForEach-Object {
                            Write-Host "Removing user: $($_.Name)"
                            try { Remove-LocalUser -Name $_.Name -ErrorAction Stop } catch { Write-Warning $_.Exception.Message }
                        }
                        
                        Write-Host "--- Removing local user profiles ---"
                        Get-ChildItem -Path 'C:\\Users' -Directory | Where-Object { $_.Name -notin $systemProfiles } | ForEach-Object {
                            $folder = $_
                            Write-Host "Removing profile folder: $($folder.FullName)"
                            try {
                                $profile = Get-CimInstance Win32_UserProfile | Where-Object { $_.LocalPath -eq $folder.FullName }
                                if ($profile) {
                                    Remove-CimInstance -InputObject $profile -ErrorAction SilentlyContinue
                                    Write-Host "Removed registry entry for: $($folder.Name)"
                                }
                                Remove-Item -Path $folder.FullName -Recurse -Force -ErrorAction Stop
                            } catch {
                                Write-Warning "Failed to remove profile '$($folder.Name)': $($_.Exception.Message)"
                            }
                        }
                    """
                )
            else:
                self.log("INFO: Skipping User Cleanup as requested.")

            # STEP 2: Remove Problematic Applications
            self.run_powershell(
                title="STEP 2: Remove Problematic Applications",
                command="""
                    function Uninstall-App {
                        param([string]$AppNamePattern)
                        $regPaths = 'HKLM:\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*', 'HKLM:\\SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*'
                        Get-ItemProperty $regPaths -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*$AppNamePattern*" -and $_.UninstallString } | ForEach-Object {
                            $app = $_
                            Write-Host "Uninstalling: $($app.DisplayName)"
                            if ($app.UninstallString -like "*msiexec*") {
                                Start-Process "msiexec.exe" -ArgumentList "/x", $app.ProductCode, "/quiet", "/norestart" -Wait
                            } else {
                                $args = $app.UninstallString.Split(' ')
                                $exe = $args[0]
                                $rest = $args[1..($args.Length-1)] + "/S", "/silent", "/quiet"
                                Start-Process $exe -ArgumentList $rest -Wait
                            }
                        }
                    }
                    Uninstall-App -AppNamePattern "Veeam"
                    Uninstall-App -AppNamePattern "SnapAgent"
                    Uninstall-App -AppNamePattern "Blackpoint"
                """
            )

            # STEP 3: Clear Agent Identity Data
            if not self.skip_agent_cleanup_var.get():
                self.run_powershell(
                    title="STEP 3: Clear Agent Identity Data",
                    command="""
                        try { Remove-ItemProperty -Path 'HKLM:\\SOFTWARE\\WOW6432Node\\NinjaRMM LLC\\NinjaRMMAgent\\Agent' -Name 'NodeId' -Force -ErrorAction Stop; Write-Host 'Removed NinjaRMM NodeId' } catch {}
                        try { Remove-Item -Path 'HKLM:\\SOFTWARE\\Veeam' -Recurse -Force -ErrorAction Stop; Write-Host 'Removed Veeam registry key' } catch {}
                        try { Remove-Item -Path 'C:\\ProgramData\\NinjaRMMAgent' -Recurse -Force -ErrorAction Stop; Write-Host 'Removed NinjaRMM data folder' } catch {}
                        try { Remove-Item -Path 'C:\\ProgramData\\Veeam' -Recurse -Force -ErrorAction Stop; Write-Host 'Removed Veeam data folder' } catch {}
                    """
                )
            else:
                self.log("INFO: Skipping Agent Cleanup as requested.")
            
            # STEP 4: Disable BitLocker
            self.run_powershell(
                title="STEP 4: Disable BitLocker",
                command="""
                    $encryptedVolumes = Get-BitLockerVolume | Where-Object { $_.VolumeStatus -ne 'FullyDecrypted' }
                    if ($encryptedVolumes) {
                        Write-Host "Disabling BitLocker on encrypted volumes..."
                        $encryptedVolumes | ForEach-Object {
                            Write-Host "Disabling on $($_.MountPoint)..."
                            try { Disable-BitLocker -MountPoint $_.MountPoint -ErrorAction Stop } catch { Write-Warning "Failed to disable BitLocker on $($_.MountPoint): $($_.Exception.Message)" }
                        }
                        # Wait for decryption
                        while ((Get-BitLockerVolume | Where-Object { $_.VolumeStatus -eq 'DecryptionInProgress' }).Count -gt 0) {
                           Write-Host "Waiting for BitLocker decryption to complete..."
                           Start-Sleep -Seconds 10
                        }
                        Write-Host "BitLocker decryption complete."
                    } else {
                        Write-Host "No encrypted BitLocker volumes found."
                    }
                """
            )

            # STEP 5: Clear Windows Logs
            if not self.skip_log_cleanup_var.get():
                self.run_powershell(
                    title="STEP 5: Clear Windows Logs",
                    command="""
                        wevtutil.exe cl Application
                        wevtutil.exe cl Security
                        wevtutil.exe cl Setup
                        wevtutil.exe cl System
                        if (Test-Path "$env:SystemRoot\\Panther") { Remove-Item -Path "$env:SystemRoot\\Panther\\*" -Recurse -Force }
                    """
                )
            else:
                self.log("INFO: Skipping Log Cleanup as requested.")

            # STEP 6: Create Unattend.xml
            self.log("INFO: STEP 6: Creating Unattend.xml for Sysprep...")
            unattend_content = """<?xml version="1.0" encoding="utf-8"?>
<unattend xmlns="urn:schemas-microsoft-com:unattend">
    <settings pass="specialize">
        <component name="Microsoft-Windows-Shell-Setup" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
            <CopyProfile>true</CopyProfile>
            <UserAccounts>
                <LocalAccounts>
                    <LocalAccount wcm:action="add" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State">
                        <Name>installadmin</Name>
                        <Group>Administrators</Group>
                        <Password><Value>DTC@dental2025</Value><PlainText>true</PlainText></Password>
                    </LocalAccount>
                </LocalAccounts>
            </UserAccounts>
            <AutoLogon>
                <Enabled>true</Enabled>
                <Username>installadmin</Username>
                <Password><Value>DTC@dental2025</Value><PlainText>true</PlainText></Password>
            </AutoLogon>
        </component>
    </settings>
    <settings pass="oobeSystem">
        <component name="Microsoft-Windows-Shell-Setup" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS">
            <OOBE><HideEULAPage>true</HideEULAPage><HideLocalAccountScreen>true</HideLocalAccountScreen><HideOnlineAccountScreens>true</HideOnlineAccountScreens><HideWirelessSetupInOOBE>true</HideWirelessSetupInOOBE><ProtectYourPC>1</ProtectYourPC><SkipMachineOOBE>true</SkipMachineOOBE></OOBE>
        </component>
    </settings>
</unattend>
"""
            sysprep_dir = Path("C:/Windows/System32/Sysprep")
            unattend_path = sysprep_dir / "unattend.xml"
            try:
                with open(unattend_path, "w", encoding="utf-8") as f:
                    f.write(unattend_content)
                self.log("SUCCESS: Created Unattend.xml")
            except Exception as e:
                self.log(f"ERROR: Failed to create Unattend.xml: {e}")
                return # Stop if we can't create this file

            # STEP 7: Run Sysprep Loop
            self.log("INFO: STEP 7: Starting Sysprep Generalization Loop...")
            max_attempts = 10
            for attempt in range(1, max_attempts + 1):
                self.log(f"--- Sysprep Attempt #{attempt} of {max_attempts} ---")
                
                # Clear Panther logs before running
                self.run_powershell(title="Clearing Sysprep Logs", command="if (Test-Path 'C:\\Windows\\System32\\Sysprep\\Panther') { Remove-Item -Path 'C:\\Windows\\System32\\Sysprep\\Panther\\*' -Recurse -Force }")

                # Run Sysprep
                sysprep_proc = subprocess.run([
                    "C:\\Windows\\System32\\Sysprep\\sysprep.exe",
                    "/generalize", "/oobe", "/shutdown", f"/unattend:{unattend_path}"
                ], capture_output=True, text=True)

                # A successful run shuts down the PC, so if we're here, it failed.
                self.log("WARNING: Sysprep process completed without shutting down, indicating a failure.")
                time.sleep(5) # Give logs time to be written

                # Check for AppX blockers
                blockers_result = subprocess.run(
                    ["powershell", "-Command", "Get-Content 'C:\\Windows\\System32\\Sysprep\\Panther\\setuperr.log' -Raw | Select-String -Pattern 'SYSPRP Package (.*?) was installed for a user' -AllMatches | ForEach-Object { $_.Matches.Groups[1].Value }"],
                    capture_output=True, text=True
                )
                
                blockers = list(set(blockers_result.stdout.strip().splitlines()))
                if not blockers:
                    self.log("ERROR: Sysprep failed, but no AppX blockers were found in the log. Manual investigation required.")
                    self.log(f"Sysprep output:\n{sysprep_proc.stdout}\n{sysprep_proc.stderr}")
                    break # Exit the loop

                self.log(f"WARNING: Found {len(blockers)} potential AppX blockers.")

                # Remove blockers
                blocker_list_str = ", ".join([f"'{b}'" for b in blockers])
                removal_command = f"""
                    $blockersToRemove = @({blocker_list_str})
                    $totalRemoved = 0
                    foreach ($blocker in $blockersToRemove) {{
                        $removed = $false
                        Get-AppxPackage -AllUsers -Name "*$blocker*" | Remove-AppxPackage -AllUsers
                        if ($?) {{ $removed = $true }}
                        Get-AppxProvisionedPackage -Online | Where-Object {{ $_.PackageName -like "*$blocker*" }} | Remove-AppxProvisionedPackage -Online
                        if ($?) {{ $removed = $true }}
                        if ($removed) {{
                             Write-Host "Removed blocker: $blocker"
                             $totalRemoved++
                        }}
                    }}
                    return $totalRemoved
                """
                removal_success = self.run_powershell(title=f"Removing {len(blockers)} blockers", command=removal_command)

                if not removal_success:
                    self.log("ERROR: Failed to remove AppX blockers. Aborting.")
                    break
                
                if attempt == max_attempts:
                    self.log("ERROR: Reached max Sysprep attempts. Aborting.")
                    break
                
                self.log("INFO: Retrying Sysprep...")
            
        except Exception as e:
            self.log(f"FATAL: The generalization process failed: {e}")
        finally:
            self.generalize_button.config(state="normal")
            self.create_button.config(state="normal")

def main():
    # Check platform first
    if not check_platform():
        return
    
    # Check if running as administrator
    import ctypes
    try:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin()
    except Exception as e:
        # Log this error to console if possible, useful for debugging
        print(f"Admin check failed: {e}")
        is_admin = False

    if not is_admin:
        # Using tk._default_root to show messagebox without a full Tk window
        # This is a bit of a hack, but works for this purpose.
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Administrator Required", "This application must be started with Administrator privileges.")
        root.destroy()
    else:
        root = tk.Tk()
        app = WindowsImagePrepGUI(root)
        root.mainloop()

if __name__ == "__main__":
    main() 