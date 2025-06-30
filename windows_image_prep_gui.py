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
import webbrowser

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
        self.root.title("OS Imaging and Processing Tool")
        self.root.geometry("750x650")
        self.root.minsize(700, 600)

        # Style
        self.style = ttk.Style(self.root)
        self.style.theme_use('vista')
        
        # --- Main UI Structure ---
        # Create a notebook (tabbed interface)
        notebook = ttk.Notebook(self.root)
        notebook.pack(pady=10, padx=10, fill="both", expand=True)

        # Create frames for each tab
        create_frame = ttk.Frame(notebook, padding="10")
        process_frame = ttk.Frame(notebook, padding="10")

        notebook.add(create_frame, text='Create VHDX Image')
        notebook.add(process_frame, text='Post-Process VHDX')

        # --- Populate Tabs ---
        self.populate_create_tab(create_frame)
        self.populate_process_tab(process_frame)

        # --- Log Area (shared across tabs) ---
        log_frame = ttk.LabelFrame(self.root, text="Log", padding="10")
        log_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.log_area = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, state='disabled', bg="#f0f0f0")
        self.log_area.pack(fill="both", expand=True)
        
        # --- Initial Checks ---
        self.check_admin()
        self.check_and_download_disk2vhd()

    def populate_create_tab(self, parent_frame):
        """Populates the 'Create VHDX Image' tab with widgets."""
        # --- Configuration Section ---
        config_frame = ttk.LabelFrame(parent_frame, text="Configuration", padding="10")
        config_frame.pack(fill="x", expand=False)
        config_frame.columnconfigure(1, weight=1)

        # Disk2vhd Path
        ttk.Label(config_frame, text="Disk2vhd Path:").grid(row=0, column=0, sticky="w", pady=2)
        self.disk2vhd_path_var = tk.StringVar(value=str(Path("./disk2vhd.exe").resolve()))
        self.disk2vhd_entry = ttk.Entry(config_frame, textvariable=self.disk2vhd_path_var, state='readonly')
        self.disk2vhd_entry.grid(row=0, column=1, sticky="we", padx=5)
        self.download_button = ttk.Button(config_frame, text="Check / Download", command=self.check_and_download_disk2vhd)
        self.download_button.grid(row=0, column=2, sticky="e")

        # Destination Path (now supports both UNC and local paths)
        ttk.Label(config_frame, text="VHDX Destination:").grid(row=1, column=0, sticky="w", pady=2)
        self.destination_path_var = tk.StringVar()
        self.destination_entry = ttk.Entry(config_frame, textvariable=self.destination_path_var)
        self.destination_entry.grid(row=1, column=1, sticky="we", padx=5)
        self.browse_destination_button = ttk.Button(config_frame, text="Browse...", command=self.browse_destination_path)
        self.browse_destination_button.grid(row=1, column=2, sticky="e")
        self.destination_entry.insert(0, "C:\\Images\\image-name.vhdx or \\\\server\\share\\image-name.vhdx")
        
        # Credentials (only needed for UNC paths)
        creds_frame = ttk.LabelFrame(config_frame, text="Network Credentials (UNC paths only)", padding="5")
        creds_frame.grid(row=2, column=0, columnspan=3, sticky="we", pady=5)
        creds_frame.columnconfigure(1, weight=1)
        
        ttk.Label(creds_frame, text="Username:").grid(row=0, column=0, sticky="w", pady=2)
        self.username_var = tk.StringVar()
        self.username_entry = ttk.Entry(creds_frame, textvariable=self.username_var)
        self.username_entry.grid(row=0, column=1, sticky="we", padx=5)
        
        ttk.Label(creds_frame, text="Password:").grid(row=1, column=0, sticky="w", pady=2)
        self.password_var = tk.StringVar()
        self.password_entry = ttk.Entry(creds_frame, textvariable=self.password_var, show="*")
        self.password_entry.grid(row=1, column=1, sticky="we", padx=5)

        # Capture Option
        self.capture_os_only_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(config_frame, text="Capture only Windows (OS) volume (Recommended)", variable=self.capture_os_only_var).grid(row=3, column=0, columnspan=3, sticky="w", pady=5)

        # --- Action Button ---
        self.create_button = ttk.Button(parent_frame, text="Create VHDX Image", command=self.start_image_creation_thread)
        self.create_button.pack(pady=10, fill="x")
        
        # --- Generalization Section ---
        generalize_frame = ttk.LabelFrame(parent_frame, text="Generalization (Sysprep)", padding="10")
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

    def populate_process_tab(self, parent_frame):
        """Populates the 'Post-Process VHDX' tab with widgets."""
        process_config_frame = ttk.LabelFrame(parent_frame, text="VHDX UEFI/GPT Conversion", padding="10")
        process_config_frame.pack(fill="x", expand=False, pady=(0, 10))
        process_config_frame.columnconfigure(1, weight=1)
        
        # VHDX Path
        ttk.Label(process_config_frame, text="VHDX Path:").grid(row=0, column=0, sticky="w", pady=2)
        self.vhdx_path_var = tk.StringVar()
        self.vhdx_path_entry = ttk.Entry(process_config_frame, textvariable=self.vhdx_path_var)
        self.vhdx_path_entry.grid(row=0, column=1, sticky="we", padx=5)
        self.vhdx_browse_button = ttk.Button(process_config_frame, text="Browse...", command=self.browse_vhdx_file)
        self.vhdx_browse_button.grid(row=0, column=2, sticky="e")
        
        # gptgen Path
        ttk.Label(process_config_frame, text="gptgen Path:").grid(row=1, column=0, sticky="w", pady=2)
        self.gptgen_path_var = tk.StringVar(value=str(Path("./gptgen.exe").resolve()))
        self.gptgen_entry = ttk.Entry(process_config_frame, textvariable=self.gptgen_path_var, state='readonly')
        self.gptgen_entry.grid(row=1, column=1, sticky="we", padx=5)
        self.gptgen_find_button = ttk.Button(process_config_frame, text="Help Me Find It", command=self.open_gptgen_download_page)
        self.gptgen_find_button.grid(row=1, column=2, sticky="e")
        
        self.process_button = ttk.Button(parent_frame, text="Process VHDX (Mount, Convert, Partition)", command=self.start_vhdx_processing_thread)
        self.process_button.pack(pady=10, fill="x")

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

    def browse_vhdx_file(self):
        """Opens a file dialog to select a VHDX file."""
        path = filedialog.askopenfilename(
            title="Select VHDX File",
            filetypes=(("VHDX Files", "*.vhdx"), ("All files", "*.*"))
        )
        if path:
            self.vhdx_path_var.set(path)

    def browse_destination_path(self):
        """Opens a file dialog to select a destination path for the VHDX file."""
        path = filedialog.asksaveasfilename(
            title="Save VHDX As...",
            defaultextension=".vhdx",
            filetypes=(("VHDX Files", "*.vhdx"), ("All files", "*.*"))
        )
        if path:
            self.destination_path_var.set(path)

    def open_gptgen_download_page(self):
        """Opens a web browser to the gptgen download page."""
        url = "https://sourceforge.net/projects/gptgen/"
        self.log(f"INFO: Opening web browser to {url} for gptgen.")
        self.log("INFO: Please download the tool, extract it, and place 'gptgen.exe' in the same folder as this application.")
        webbrowser.open(url)

    def start_image_creation_thread(self):
        """Starts the imaging process in a new thread to avoid freezing the GUI."""
        self.create_button.config(state="disabled")
        thread = threading.Thread(target=self.create_image_worker)
        thread.daemon = True
        thread.start()

    def create_image_worker(self):
        """The main worker function that performs all imaging tasks."""
        share_path = None # For cleanup
        try:
            # 1. Get parameters from GUI
            destination_path = self.destination_path_var.get().strip()
            username = self.username_var.get()
            password = self.password_var.get()
            disk2vhd_exe = self.disk2vhd_path_var.get()

            # Validate destination path
            if not destination_path:
                self.log("ERROR: No destination path provided.")
                messagebox.showerror("Invalid Input", "Please provide a destination path for the VHDX file.")
                return

            # Clear placeholder text if it's still there
            if "or \\\\" in destination_path:
                self.log("ERROR: Please replace the placeholder text with an actual path.")
                messagebox.showerror("Invalid Input", "Please provide a valid destination path (local or UNC).")
                return

            is_unc_path = destination_path.startswith("\\\\")
            
            if is_unc_path:
                self.log(f"INFO: Using UNC path: {destination_path}")
            else:
                self.log(f"INFO: Using local path: {destination_path}")
                # Validate local path - check if directory exists
                local_dir = Path(destination_path).parent
                if not local_dir.exists():
                    self.log(f"ERROR: Local directory does not exist: {local_dir}")
                    if messagebox.askyesno("Create Directory", f"The directory {local_dir} does not exist. Create it?"):
                        try:
                            local_dir.mkdir(parents=True, exist_ok=True)
                            self.log(f"SUCCESS: Created directory: {local_dir}")
                        except Exception as e:
                            self.log(f"ERROR: Failed to create directory: {e}")
                            return
                    else:
                        return

            # 2. Find OS volumes
            self.log("INFO: Identifying OS volumes to capture...")
            try:
                if self.capture_os_only_var.get():
                    self.log("INFO: 'Capture only OS volume' is selected.")
                    system_drive = os.environ.get('SystemDrive')
                    if not system_drive:
                        raise ValueError("Could not determine SystemDrive from environment variables.")
                    volumes_to_capture = [system_drive]
                else:
                    self.log("INFO: Capturing all volumes on the OS disk.")
                    ps_command_disk = "(Get-Partition | Where-Object { $_.DriveLetter -eq $env:SystemDrive.Trim(':') }).DiskNumber"
                    disk_number_result = subprocess.run(["powershell", "-Command", ps_command_disk], capture_output=True, text=True, check=True, encoding='utf-8', errors='ignore')
                    disk_number = disk_number_result.stdout.strip()

                    ps_command_volumes = f"(Get-Partition -DiskNumber {disk_number} | ForEach-Object {{ $_.DriveLetter }} | Where-Object {{ -not [string]::IsNullOrWhiteSpace($_) }})"
                    volumes_result = subprocess.run(["powershell", "-Command", ps_command_volumes], capture_output=True, text=True, check=True, encoding='utf-8', errors='ignore')
                    volumes_to_capture = [v.strip() + ":" for v in volumes_result.stdout.strip().splitlines() if v.strip()]
                
                if not volumes_to_capture:
                    raise ValueError("Could not auto-detect any volumes.")

                self.log(f"INFO: Volumes to be captured: {', '.join(volumes_to_capture)}")
            except (subprocess.CalledProcessError, ValueError) as e:
                self.log(f"ERROR: Failed to identify OS volumes: {e}")
                return
            
            # 3. Authenticate to network share if it's a UNC path and credentials are provided
            if is_unc_path and username and password:
                destination_path_obj = Path(destination_path)
                share_path = str(destination_path_obj.parent)
                self.log(f"INFO: Authenticating to network share '{share_path}'...")
                net_use_cmd = ["net", "use", share_path, password, f"/user:{username}"]
                
                map_proc = subprocess.run(net_use_cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
                if map_proc.returncode != 0:
                    self.log(f"ERROR: Failed to authenticate to network share. {map_proc.stderr or map_proc.stdout}")
                    return
                self.log("SUCCESS: Authenticated to network share successfully.")
            elif is_unc_path and not (username and password):
                self.log("WARNING: UNC path specified but no credentials provided. Attempting to access with current user credentials.")

            # 4. Run Disk2vhd
            self.log("INFO: Starting Disk2vhd capture. This may take a long time...")
            # Using /accepteula is a common requirement for sysinternals tools.
            # The .vhdx extension on the output file is usually sufficient.
            arguments = ["/accepteula"] + volumes_to_capture + [destination_path]
            command_to_run = [disk2vhd_exe] + arguments
            self.log(f"COMMAND: \"{disk2vhd_exe}\" {' '.join(arguments)}")

            process = subprocess.Popen(
                ' '.join(command_to_run), 
                stdout=subprocess.PIPE, 
                stderr=subprocess.STDOUT, 
                text=True, 
                encoding='utf-8',
                errors='ignore',
                shell=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            
            if process.stdout:
                for line in iter(process.stdout.readline, ''):
                    self.log(line.strip())
            
            process.wait()

            if process.returncode == 0:
                self.log("SUCCESS: Disk2vhd completed successfully!")
                self.log(f"SUCCESS: Image saved to: {destination_path}")
            else:
                self.log(f"ERROR: Disk2vhd failed with exit code: {process.returncode}.")

        except Exception as e:
            self.log(f"FATAL: An unexpected error occurred: {e}")
        finally:
            # 5. Cleanup network connection if used
            if share_path:
                self.log(f"INFO: Cleaning up network connection to '{share_path}'...")
                unmap_proc = subprocess.run(["net", "use", share_path, "/delete"], capture_output=True, text=True)
                if unmap_proc.returncode == 0:
                    self.log(f"INFO: Network connection to '{share_path}' has been removed.")
                else:
                    self.log(f"WARN: Failed to remove network connection to '{share_path}'. {unmap_proc.stderr or unmap_proc.stdout}")
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
                encoding='utf-8',
                errors='ignore',
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
                    capture_output=True, text=True, encoding='utf-8', errors='ignore'
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

    def start_vhdx_processing_thread(self):
        """Starts the VHDX processing in a new thread."""
        vhdx_path = self.vhdx_path_var.get()
        if not vhdx_path or not Path(vhdx_path).exists():
            messagebox.showerror("Invalid Path", "Please select a valid VHDX file.")
            return

        gptgen_path = Path(self.gptgen_path_var.get())
        if not gptgen_path.exists():
            messagebox.showerror("gptgen Not Found", f"gptgen.exe not found at {gptgen_path}.\nPlease use the 'Help Me Find It' button to download and place it correctly.")
            return

        if not messagebox.askyesno("Confirm VHDX Processing", 
                                   f"This will modify the VHDX at:\n{vhdx_path}\n\nThe process involves mounting, converting to GPT, and re-partitioning. This is a destructive operation on the VHDX file. Are you sure?"):
            return
            
        self.process_button.config(state="disabled")
        thread = threading.Thread(target=self.vhdx_processing_worker, args=(vhdx_path, str(gptgen_path)))
        thread.daemon = True
        thread.start()

    def vhdx_processing_worker(self, vhdx_path, gptgen_path):
        """Mounts, converts, and partitions the VHDX."""
        disk_number = None
        try:
            self.log("--- Starting VHDX Post-Processing ---")
            
            # 1. Mount VHDX
            self.log(f"INFO: Mounting VHDX: {vhdx_path}")
            mount_command = f"""
                $vhdx = Mount-Vhd -Path '{vhdx_path}' -Passthru
                $diskNumber = ($vhdx | Get-Disk).DiskNumber
                Write-Host "VHDX mounted as Disk $diskNumber"
                return $diskNumber
            """
            # Using run instead of run_powershell to capture output directly
            proc = subprocess.run(["powershell", "-Command", mount_command], capture_output=True, text=True, encoding='utf-8', errors='ignore')
            if proc.returncode != 0:
                self.log(f"ERROR: Failed to mount VHDX. PowerShell output:\n{proc.stdout}\n{proc.stderr}")
                return
            
            # Extract disk number from output
            match = re.search(r'Disk (\d+)', proc.stdout)
            if not match:
                self.log(f"ERROR: Could not determine disk number for mounted VHDX. Output:\n{proc.stdout}")
                # Attempt to find it another way before giving up
                ps_find_disk = f"($vhd = Get-Vhd -Path '{vhdx_path}'; if($vhd) {{ ($vhd | Get-Disk).DiskNumber }} else {{ -1 }})"
                disk_number_res = subprocess.run(['powershell', '-Command', ps_find_disk], capture_output=True, text=True)
                disk_number = disk_number_res.stdout.strip()
                if disk_number == "-1" or not disk_number.isdigit():
                    self.log("ERROR: Secondary check also failed. Aborting.")
                    return
                self.log(f"INFO: Found disk number via secondary method: {disk_number}")
            else:
                disk_number = match.group(1)
            
            self.log(f"SUCCESS: VHDX is mounted as Disk {disk_number}.")

            # 2. Convert to GPT
            self.log(f"INFO: Converting Disk {disk_number} to GPT using gptgen...")
            gptgen_proc = subprocess.run([gptgen_path, "-w", f"\\\\.\\physicaldrive{disk_number}"], capture_output=True, text=True)
            if gptgen_proc.returncode != 0:
                self.log(f"ERROR: gptgen failed. Output:\n{gptgen_proc.stdout}\n{gptgen_proc.stderr}")
                return
            self.log("SUCCESS: gptgen conversion completed.")

            # 3. Repartition using diskpart
            self.log("INFO: Re-partitioning the disk to create EFI and Recovery partitions...")
            diskpart_script = f"""
select disk {disk_number}
select partition 1
shrink desired=8192 minimum=8192
create partition efi size=4096
format fs=fat32 quick label="EFI"
assign letter="S"
create partition primary size=4096
format fs=ntfs quick label="Recovery"
set id="de94bba4-06d1-4d40-a16a-bfd50179d6ac"
gpt attributes=0x8000000000000001
"""
            script_path = Path("./diskpart_script.txt")
            with open(script_path, "w") as f:
                f.write(diskpart_script)
            
            diskpart_proc = subprocess.run(["diskpart", "/s", str(script_path)], capture_output=True, text=True)
            script_path.unlink() # Clean up script file
            if diskpart_proc.returncode != 0:
                 self.log(f"ERROR: diskpart failed. Output:\n{diskpart_proc.stdout}\n{diskpart_proc.stderr}")
                 return
            self.log("SUCCESS: Disk re-partitioned.")

            # 4. Rebuild boot data
            self.log("INFO: Rebuilding boot data on new EFI partition (S:)...")
            # Find the Windows directory. It should be the largest partition, likely mounted.
            ps_get_win_drive = f"Get-Partition -DiskNumber {disk_number} | Where-Object {{ $_.Type -eq 'Basic' }} | Sort-Object -Property Size -Descending | Select-Object -First 1 | Select-Object -ExpandProperty DriveLetter"
            win_drive_res = subprocess.run(['powershell', '-Command', ps_get_win_drive], capture_output=True, text=True)
            win_drive_letter = win_drive_res.stdout.strip()
            if not win_drive_letter:
                self.log("ERROR: Could not auto-detect the Windows drive letter on the mounted VHDX. Aborting bcdboot.")
                return

            self.log(f"INFO: Found Windows partition at {win_drive_letter}:")
            win_path = f"{win_drive_letter}:\\Windows"
            
            bcdboot_proc = subprocess.run(["bcdboot", win_path, "/s", "S:", "/f", "UEFI"], capture_output=True, text=True)
            if bcdboot_proc.returncode != 0:
                self.log(f"ERROR: bcdboot failed. Output:\n{bcdboot_proc.stdout}\n{bcdboot_proc.stderr}")
                return
            self.log("SUCCESS: Boot data rebuilt.")

            self.log("--- VHDX Post-Processing COMPLETED ---")

        except Exception as e:
            self.log(f"FATAL: An unexpected error occurred during VHDX processing: {e}")
        finally:
            # 5. Cleanup: Dismount VHDX
            self.log("INFO: Cleaning up by dismounting the VHDX...")
            dismount_command = f"Dismount-Vhd -Path '{vhdx_path}'"
            if disk_number:
                 dismount_command = f"Dismount-Vhd -DiskNumber {disk_number}"
            
            dismount_proc = subprocess.run(["powershell", "-Command", dismount_command], capture_output=True, text=True)
            if dismount_proc.returncode == 0:
                self.log("SUCCESS: VHDX dismounted.")
            else:
                self.log(f"WARN: Failed to dismount VHDX. It may need to be dismounted manually via Disk Management. Error:\n{dismount_proc.stderr or dismount_proc.stdout}")
            
            self.process_button.config(state="normal")

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