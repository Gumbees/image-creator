#!/usr/bin/env python3
"""
Windows Image Preparation GUI
A comprehensive tool for preparing Windows images for generalization
"""

import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog, simpledialog
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
import shutil
import sys
import winreg
import sqlite3
import json
import uuid
from datetime import datetime
import math

def check_platform():
    """Check if running on Windows"""
    if platform.system() != 'Windows':
        print(f"Error: This tool is designed for Windows systems only.")
        print(f"Current platform: {platform.system()}")
        print("Please run this tool on a Windows machine where you're preparing the image.")
        return False
    return True

def generate_uuidv7():
    """Generate a UUIDv7 (time-ordered UUID)"""
    # UUIDv7 implementation - timestamp-based
    timestamp_ms = int(time.time() * 1000)
    
    # 48-bit timestamp (milliseconds since epoch)
    time_high = (timestamp_ms >> 16) & 0xFFFFFFFF
    time_low = timestamp_ms & 0xFFFF
    
    # 12-bit random data + version
    rand_a = (uuid.uuid4().int >> 76) & 0xFFF
    version = 7
    
    # 62-bit random data + variant
    rand_b = uuid.uuid4().int & 0x3FFFFFFFFFFFFFFF
    variant = 0x8000000000000000
    
    # Construct UUID
    uuid_int = (time_high << 96) | (time_low << 80) | (version << 76) | (rand_a << 64) | variant | rand_b
    return str(uuid.UUID(int=uuid_int))

class DatabaseManager:
    """Manages SQLite database for image management"""
    
    def __init__(self, workflow_mode=None):
        # Determine database path based on workflow mode
        if workflow_mode == "development":
            # Development mode: use temp folder (cleared on generalization/cleanup)
            windir = os.environ.get('WINDIR', 'C:\\Windows')
            self.db_path = Path(windir) / "Temp" / "pyc.db"
        else:
            # Production mode or unknown: use permanent location
            self.db_path = Path(os.environ.get('PUBLIC', 'C:\\Users\\Public')) / "Documents" / "pyc" / "pyc.db"
        
        self.config_initialized = False
        self.init_database()
    
    def init_database(self):
        """Initialize database and create tables if needed"""
        # Ensure directory exists
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            
            # Config table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS config (
                    key TEXT PRIMARY KEY,
                    value TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            
            # Clients table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS clients (
                    id TEXT PRIMARY KEY,
                    name TEXT NOT NULL,
                    short_name TEXT NOT NULL UNIQUE,
                    description TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            
            # Sites table (belong to clients)
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS sites (
                    id TEXT PRIMARY KEY,
                    client_id TEXT NOT NULL,
                    name TEXT NOT NULL,
                    short_name TEXT NOT NULL,
                    description TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (client_id) REFERENCES clients (id),
                    UNIQUE(client_id, short_name)
                )
            ''')
            
            # Images table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS images (
                    id TEXT PRIMARY KEY,
                    client_id TEXT NOT NULL,
                    site_id TEXT NOT NULL,
                    role TEXT NOT NULL,
                    repository_path TEXT NOT NULL,
                    repository_size_gb INTEGER DEFAULT 0,
                    snapshot_count INTEGER DEFAULT 0,
                    latest_snapshot_id TEXT,
                    restic_password TEXT,
                    status TEXT DEFAULT 'ready',
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (client_id) REFERENCES clients (id),
                    FOREIGN KEY (site_id) REFERENCES sites (id)
                )
            ''')
            
            # Database migrations - add missing columns if they don't exist
            self.migrate_database(cursor)
            
            conn.commit()
            
            # Check if this is first launch
            cursor.execute('SELECT COUNT(*) FROM config')
            if cursor.fetchone()[0] == 0:
                self.config_initialized = False
            else:
                self.config_initialized = True
    
    def migrate_database(self, cursor):
        """Apply database migrations for repository-only storage"""
        try:
            # Check which columns exist in images table
            cursor.execute("PRAGMA table_info(images)")
            columns = [column[1] for column in cursor.fetchall()]
            
            # Add required restic-specific columns if missing
            if 'repository_path' not in columns:
                cursor.execute('ALTER TABLE images ADD COLUMN repository_path TEXT')
                print("Added repository_path column to images table")
                
            if 'repository_size_gb' not in columns:
                cursor.execute('ALTER TABLE images ADD COLUMN repository_size_gb INTEGER DEFAULT 0')
                print("Added repository_size_gb column to images table")
                
            if 'snapshot_count' not in columns:
                cursor.execute('ALTER TABLE images ADD COLUMN snapshot_count INTEGER DEFAULT 0')
                print("Added snapshot_count column to images table")
                
            if 'latest_snapshot_id' not in columns:
                cursor.execute('ALTER TABLE images ADD COLUMN latest_snapshot_id TEXT')
                print("Added latest_snapshot_id column to images table")
                
            if 'restic_password' not in columns:
                cursor.execute('ALTER TABLE images ADD COLUMN restic_password TEXT')
                print("Added restic_password column to images table")
            
            # Update status column default if it exists
            if 'status' not in columns:
                cursor.execute('ALTER TABLE images ADD COLUMN status TEXT DEFAULT "ready"')
                print("Added status column to images table")
            
            # Migration: Delete any legacy VHDX entries (image_type != 'restic')
            try:
                cursor.execute("DELETE FROM images WHERE image_type IS NOT NULL AND image_type != 'restic'")
                cursor.execute("DELETE FROM images WHERE repository_path IS NULL OR repository_path = ''")
                print("Cleaned up legacy VHDX entries")
            except:
                pass  # Ignore if columns don't exist yet
                
        except Exception as e:
            print(f"Database migration warning: {e}")
    
    def get_config(self, key, default=None):
        """Get configuration value"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT value FROM config WHERE key = ?', (key,))
            result = cursor.fetchone()
            return result[0] if result else default
    
    def set_config(self, key, value):
        """Set configuration value"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                INSERT OR REPLACE INTO config (key, value, updated_at) 
                VALUES (?, ?, CURRENT_TIMESTAMP)
            ''', (key, value))
            conn.commit()
    
    def get_s3_config(self):
        """Get S3 configuration from database"""
        s3_config = {}
        for key in ['s3_bucket', 's3_access_key', 's3_secret_key', 's3_endpoint', 's3_region']:
            value = self.get_config(key)
            if value:
                s3_config[key] = value
        return s3_config if s3_config else None
    
    def set_s3_config(self, bucket, access_key, secret_key, endpoint, region="us-east-1"):
        """Set S3 configuration in database"""
        self.set_config('s3_bucket', bucket)
        self.set_config('s3_access_key', access_key)
        self.set_config('s3_secret_key', secret_key)
        self.set_config('s3_endpoint', endpoint)
        self.set_config('s3_region', region)
            
    def get_working_vhdx_directory(self):
        """Get or create working VHDX directory on largest volume"""
        # Check if already configured
        working_dir = self.get_config('working_vhdx_directory')
        if working_dir and Path(working_dir).exists():
            return Path(working_dir)
        
        # Find largest volume and create working directory
        largest_volume = self.find_largest_volume()
        working_dir = largest_volume / "pyc-working-vhdx"
        working_dir.mkdir(parents=True, exist_ok=True)
        
        # Store in config
        self.set_config('working_vhdx_directory', str(working_dir))
        return working_dir
        
    def find_largest_volume(self):
        """Find the volume with the most free space"""
        try:
            import shutil
            volumes = []
            
            # Get all available drives
            for drive_letter in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':
                drive_path = f"{drive_letter}:\\"
                if os.path.exists(drive_path):
                    try:
                        free_space = shutil.disk_usage(drive_path).free
                        volumes.append((drive_path, free_space))
                    except:
                        continue
            
            if volumes:
                # Return the drive with most free space
                largest_drive = max(volumes, key=lambda x: x[1])[0]
                return Path(largest_drive)
            else:
                # Fallback to C: drive
                return Path("C:\\")
                
        except Exception as e:
            print(f"Error finding largest volume: {e}")
            return Path("C:\\")
    
    def add_client(self, name, short_name, description=""):
        """Add new client"""
        client_id = generate_uuidv7()
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO clients (id, name, short_name, description)
                VALUES (?, ?, ?, ?)
            ''', (client_id, name, short_name, description))
            conn.commit()
        return client_id
    
    def add_site(self, client_id, name, short_name, description=""):
        """Add new site for client"""
        site_id = generate_uuidv7()
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO sites (id, client_id, name, short_name, description)
                VALUES (?, ?, ?, ?, ?)
            ''', (site_id, client_id, name, short_name, description))
            conn.commit()
        return site_id
    
    def get_clients(self):
        """Get all clients"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT id, name, short_name, description FROM clients ORDER BY name')
            return cursor.fetchall()
    
    def get_sites(self, client_id=None):
        """Get sites, optionally filtered by client"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            if client_id:
                cursor.execute('''
                    SELECT s.id, s.client_id, s.name, s.short_name, s.description, c.name as client_name
                    FROM sites s 
                    JOIN clients c ON s.client_id = c.id 
                    WHERE s.client_id = ? 
                    ORDER BY s.name
                ''', (client_id,))
            else:
                cursor.execute('''
                    SELECT s.id, s.client_id, s.name, s.short_name, s.description, c.name as client_name
                    FROM sites s 
                    JOIN clients c ON s.client_id = c.id 
                    ORDER BY c.name, s.name
                ''')
            return cursor.fetchall()
    
    def create_image(self, image_id, client_id, site_id, role, repository_path, repository_size_gb=0, 
                    snapshot_count=0, latest_snapshot_id=None, restic_password=None):
        """Create new restic repository image record"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO images (id, client_id, site_id, role, repository_path, repository_size_gb, 
                                  snapshot_count, latest_snapshot_id, restic_password)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (image_id, client_id, site_id, role, repository_path, repository_size_gb, 
                  snapshot_count, latest_snapshot_id, restic_password))
            conn.commit()
        return image_id
        
    def update_repository_info(self, image_id, snapshot_count=None, latest_snapshot_id=None, repository_size_gb=None):
        """Update repository information"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            updates = []
            params = []
            
            if snapshot_count is not None:
                updates.append("snapshot_count = ?")
                params.append(snapshot_count)
            if latest_snapshot_id is not None:
                updates.append("latest_snapshot_id = ?")
                params.append(latest_snapshot_id)
            if repository_size_gb is not None:
                updates.append("repository_size_gb = ?")
                params.append(repository_size_gb)
                
            if updates:
                updates.append("updated_at = CURRENT_TIMESTAMP")
                params.append(image_id)
                cursor.execute(f'''
                    UPDATE images SET {", ".join(updates)} WHERE id = ?
                ''', params)
                conn.commit()
    
    def get_client_repositories(self, client_id):
        """Get all repositories for a specific client with local file enumeration"""
        repositories = []
        
        # Get repositories from database
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT i.id, i.role, i.repository_path, i.repository_size_gb, i.snapshot_count, 
                       i.latest_snapshot_id, i.status, i.created_at, i.restic_password,
                       s.name as site_name, s.short_name as site_short
                FROM images i
                LEFT JOIN sites s ON i.site_id = s.id
                WHERE i.client_id = ? AND i.repository_path IS NOT NULL
                ORDER BY i.created_at DESC
            ''', (client_id,))
            
            db_repos = cursor.fetchall()
        
        # Add database repositories
        for repo in db_repos:
            repositories.append({
                'id': repo[0],
                'role': repo[1],
                'path': repo[2],
                'size_gb': repo[3],
                'snapshot_count': repo[4],
                'latest_snapshot': repo[5],
                'status': repo[6],
                'created_at': repo[7],
                'password': repo[8],
                'site_name': repo[9],
                'site_short': repo[10],
                'source': 'database'
            })
        
        # Also check for local repositories on disk (organized by client)
        try:
            app = WindowsImagePrepGUI.__instance if hasattr(WindowsImagePrepGUI, '__instance') else None
            if app:
                restic_base = app.get_restic_base_path()
                client_dir = restic_base / client_id
                
                if client_dir.exists():
                    for repo_dir in client_dir.iterdir():
                        if repo_dir.is_dir():
                            repo_path = str(repo_dir)
                            # Check if this repository is already in database
                            found_in_db = any(r['path'] == repo_path for r in repositories)
                            
                            if not found_in_db:
                                # Add as untracked repository
                                repositories.append({
                                    'id': None,
                                    'role': 'untracked',
                                    'path': repo_path,
                                    'size_gb': 0,
                                    'snapshot_count': 0,
                                    'latest_snapshot': None,
                                    'status': 'untracked',
                                    'created_at': None,
                                    'password': None,
                                    'site_name': 'Unknown',
                                    'site_short': 'UNK',
                                    'source': 'local_disk'
                                })
        except Exception as e:
            # Silently handle errors in local enumeration
            pass
        
        return repositories
    
    def get_images(self):
        """Get all repositories with client and site info"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT i.id, i.role, i.repository_path, i.repository_size_gb, i.snapshot_count, 
                       i.latest_snapshot_id, i.status, i.created_at,
                       c.name as client_name, c.short_name as client_short,
                       s.name as site_name, s.short_name as site_short,
                       i.restic_password
                FROM images i
                JOIN clients c ON i.client_id = c.id
                JOIN sites s ON i.site_id = s.id
                WHERE i.repository_path IS NOT NULL
                ORDER BY i.created_at DESC
            ''')
            return cursor.fetchall()
    
    def save_image_metadata(self, image_id, image_store_path):
        """Save image metadata as JSON file"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT i.*, c.name as client_name, c.short_name as client_short,
                       s.name as site_name, s.short_name as site_short
                FROM images i
                JOIN clients c ON i.client_id = c.id  
                JOIN sites s ON i.site_id = s.id
                WHERE i.id = ?
            ''', (image_id,))
            
            row = cursor.fetchone()
            if row:
                columns = [description[0] for description in cursor.description]
                metadata = dict(zip(columns, row))
                
                # Save JSON metadata file
                json_path = Path(image_store_path) / f"{image_id}.metadata.json"
                with open(json_path, 'w') as f:
                    json.dump(metadata, f, indent=2, default=str)
                
                return json_path
        return None

    def find_client_by_name(self, name):
        """Find client by name - returns (id, name, short_name, description) or None"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT id, name, short_name, description FROM clients WHERE name = ?', (name,))
            return cursor.fetchone()
    
    def find_client_by_short_name(self, short_name):
        """Find client by short name - returns (id, name, short_name, description) or None"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT id, name, short_name, description FROM clients WHERE short_name = ?', (short_name,))
            return cursor.fetchone()
    
    def get_client_site_short_names(self, client_id, site_id):
        """Get client and site short names in one query"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT c.short_name as client_short, s.short_name as site_short
                FROM clients c
                JOIN sites s ON c.id = s.client_id
                WHERE c.id = ? AND s.id = ?
            ''', (client_id, site_id))
            return cursor.fetchone()
    
    def generate_secure_password(self, client_name="", site_name="", role=""):
        """Generate a secure password for restic repository"""
        import secrets
        import string
        
        # Generate a random secure password
        alphabet = string.ascii_letters + string.digits + "!@#$%^&*"
        password = ''.join(secrets.choice(alphabet) for _ in range(16))
        
        # Create a memorable identifier for password manager
        identifier_parts = []
        if client_name:
            identifier_parts.append(client_name.replace(" ", ""))
        if site_name:
            identifier_parts.append(site_name.replace(" ", ""))
        if role:
            identifier_parts.append(role)
        
        identifier = "-".join(identifier_parts) if identifier_parts else "PYC-Restic"
        
        return password, f"PYC-{identifier}-ResticRepo"

    def get_images_by_client_and_environment(self, client_id, environment):
        """Get images filtered by client ID and environment (development/production)"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT i.id, i.client_id, i.site_id, i.role, i.created_at, 
                       i.repository_path, i.repository_size_gb, i.snapshot_count, i.latest_snapshot_id
                FROM images i
                WHERE i.client_id = ? AND i.repository_path LIKE ?
                ORDER BY i.created_at DESC
            ''', (client_id, f"%/{environment}/%"))
            return cursor.fetchall()

    def scan_s3_for_images_filtered(self, environment_filter=None):
        """Scan S3 repository for images with optional environment filtering"""
        # This method would implement S3 scanning logic
        # For now, it's a placeholder that would be implemented with AWS SDK
        # The actual implementation would:
        # 1. Connect to S3 using stored credentials
        # 2. List objects in the bucket with the specified environment filter
        # 3. Parse metadata JSON files
        # 4. Populate the database with discovered images
        pass

    def get_sites_by_client(self, client_id):
        """Get all sites for a specific client"""
        return self.get_sites(client_id)

    def get_site_by_short_name(self, short_name):
        """Find site by short name - returns (id, name, short_name, client_id) or None"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT id, name, short_name, client_id FROM sites WHERE short_name = ?', (short_name,))
            return cursor.fetchone()

    def create_client(self, client_id, name, short_name, description=""):
        """Create a new client with specified UUID"""
        return self.add_client(name, short_name, description)

    def create_site(self, site_id, name, short_name, client_id, description=""):
        """Create a new site with specified UUID"""
        return self.add_site(client_id, name, short_name, description)
    
    def add_image(self, client_id, site_id, role, wim_source_path, repository_path, repository_size_gb=0, 
                 snapshot_count=0, latest_snapshot_id=None, restic_password=None):
        """Add a new image (alias for create_image for backward compatibility)"""
        image_id = generate_uuidv7()
        return self.create_image(image_id, client_id, site_id, role, repository_path, repository_size_gb, 
                               snapshot_count, latest_snapshot_id, restic_password)
    
    def get_client_name(self, client_id):
        """Get client name by ID"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT name FROM clients WHERE id = ?', (client_id,))
            result = cursor.fetchone()
            return result[0] if result else None
    
    def get_site_name(self, site_id):
        """Get site name by ID"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT name FROM sites WHERE id = ?', (site_id,))
            result = cursor.fetchone()
            return result[0] if result else None
    
    def get_client_by_id(self, client_id):
        """Get client information by ID"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT id, name, short_name, description FROM clients WHERE id = ?', (client_id,))
            return cursor.fetchone()
    
    def get_site_by_id(self, site_id):
        """Get site information by ID"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT id, client_id, name, short_name, description FROM sites WHERE id = ?', (site_id,))
            return cursor.fetchone()
    
    @property
    def connection(self):
        """Get database connection for direct access"""
        return sqlite3.connect(self.db_path)

class WindowsImagePrepGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("OS Imaging and Processing Tool - Professional Edition")
        self.root.geometry("900x800")
        self.root.minsize(850, 750)

        # Store instance reference for database access
        WindowsImagePrepGUI.__instance = self

        # Determine workflow mode and initialize appropriate database
        workflow_mode = self.detect_workflow_mode()
        
        # Initialize database manager with workflow mode
        self.db_manager = DatabaseManager(workflow_mode)
        self.db = self.db_manager  # Keep backward compatibility
        
        # S3 configuration and workflow mode are now handled per-mode

        # Style
        self.style = ttk.Style(self.root)
        self.style.theme_use('vista')
        
        # Current step tracking
        self.current_step = 1
        self.total_steps = 3
        
        # First-time setup is now handled per-mode as needed
        
        # Get image storage path
        self.image_store_path = self.get_image_store_path()
        
        
        # --- Mode-Based UI Structure ---
        self.current_mode = None
        
        # Create main mode selection screen
        self.create_mode_selection_screen()
        
        # Initialize mode frames (created but hidden)
        self.mode_frames = {}
        
        # Initialize step frames for legacy compatibility
        self.step_frames = {}
        
        # Initialize step buttons for legacy compatibility
        self.step_buttons = {}
        
        # Initialize navigation elements for legacy compatibility
        self.prev_button = None
        self.next_button = None
        self.current_step_label = None
        
        # --- Log Area (shared across all modes) ---
        log_frame = ttk.LabelFrame(self.root, text="Process Log", padding="5")
        log_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.log_area = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, state='disabled', bg="#f0f0f0", height=8)
        self.log_area.pack(fill="both", expand=True)
        
        # --- Initial Checks ---
        self.check_admin()
        
        # --- Welcome Message for Mode-Based Interface ---
        self.log("=== Windows Image Preparation Tool - Mode-Based Interface ===")
        self.log("INFO: Welcome! Select an operating mode from the buttons above:")
        self.log("  🔧 DEVELOP CAPTURE: Create development images with S3 integration")
        self.log("  🚀 PRODUCTION CAPTURE: Create production-ready deployment images")
        self.log("  🛠️ GENERALIZE: Prepare images for deployment with sysprep")
        self.log("  📁 MANAGE IMAGES: Browse and manage existing images")
        self.log("")
        self.log("INFO: Using modern Restic backup engine with S3 cloud storage")
        self.log("="*60)

    def create_centered_dialog(self, title, width=600, height=500, resizable=True):
        """Helper method to create centered modal dialogs with consistent styling"""
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry(f"{width}x{height}")
        dialog.transient(self.root)
        dialog.grab_set()
        
        if not resizable:
            dialog.resizable(False, False)
        
        # Center dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        return dialog

    def create_top_buttons(self):
        """Create top button bar with Image Manager and other tools"""
        top_frame = ttk.Frame(self.root)
        top_frame.pack(fill="x", padx=10, pady=5)
        
        # Image Manager button
        self.image_manager_button = ttk.Button(top_frame, text="📁 Image Manager", 
                                             command=self.open_image_manager,
                                             style="Accent.TButton")
        self.image_manager_button.pack(side="left", padx=(0, 10))
        
        # Install to Public Desktop button
        self.install_button = ttk.Button(top_frame, text="📥 Install to Public Desktop", 
                                       command=self.install_to_public_desktop)
        self.install_button.pack(side="right")

    def check_for_wim_imports(self):
        """Check for WIM files in storage directory and offer to import them"""
        try:
            wim_files = list(self.image_store_path.glob("*.wim"))
            
            # Filter out files that are already in database
            images = self.db.get_images()
            known_paths = set()
            for image_data in images:
                image_path = image_data[2]  # image_path column
                if image_path:
                    known_paths.add(Path(image_path).name)
            
            orphan_wims = []
            for wim_file in wim_files:
                if wim_file.name not in known_paths:
                    orphan_wims.append(wim_file)
            
            if orphan_wims:
                self.show_wim_import_dialog(orphan_wims)
                
        except Exception as e:
            self.log(f"WARNING: Failed to check for WIM imports: {e}")

    def show_wim_import_dialog(self, wim_files):
        """Show dialog to import WIM files"""
        if not wim_files:
            return
            
        dialog = self.create_centered_dialog("Import WIM Files Found", 600, 500)
        
        ttk.Label(dialog, text="WIM Files Found for Import", 
                 font=("TkDefaultFont", 14, "bold")).pack(pady=10)
        
        ttk.Label(dialog, text=f"Found {len(wim_files)} WIM files in storage directory that are not in database:", 
                 font=("TkDefaultFont", 10)).pack(pady=5)
        
        # File list with checkboxes
        list_frame = ttk.LabelFrame(dialog, text="Select files to import", padding="10")
        list_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Scrollable frame for file list
        canvas = tk.Canvas(list_frame)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # File checkboxes
        file_vars = {}
        for i, wim_file in enumerate(wim_files):
            var = tk.BooleanVar(value=True)
            file_vars[wim_file] = var
            
            file_frame = ttk.Frame(scrollable_frame)
            file_frame.pack(fill="x", pady=2)
            
            ttk.Checkbutton(file_frame, text=wim_file.name, variable=var).pack(side="left")
            
            # Show file size
            try:
                size_gb = wim_file.stat().st_size / (1024**3)
                ttk.Label(file_frame, text=f"({size_gb:.1f} GB)", 
                         foreground="gray").pack(side="right")
            except:
                pass
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Buttons
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=20)
        
        def import_selected():
            selected_files = [f for f, var in file_vars.items() if var.get()]
            if selected_files:
                dialog.destroy()
                self.import_wim_files(selected_files)
        
        def ignore_all():
            dialog.destroy()
        
        ttk.Button(button_frame, text="Import Selected", command=import_selected).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Ignore", command=ignore_all).pack(side="left", padx=5)

    def import_wim_files(self, wim_files):
        """Import selected WIM files into database"""
        for wim_file in wim_files:
            try:
                self.import_single_wim_file(wim_file)
            except Exception as e:
                self.log(f"ERROR: Failed to import {wim_file.name}: {e}")

    def import_single_wim_file(self, wim_file):
        """Import a single WIM file with client/site selection"""
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Import: {wim_file.name}")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        ttk.Label(dialog, text=f"Import: {wim_file.name}", 
                 font=("TkDefaultFont", 12, "bold")).pack(pady=10)
        
        # Client selection
        client_frame = ttk.LabelFrame(dialog, text="Client", padding="10")
        client_frame.pack(fill="x", padx=20, pady=5)
        
        client_var = tk.StringVar()
        client_combo = ttk.Combobox(client_frame, textvariable=client_var, width=40)
        client_combo.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        ttk.Button(client_frame, text="New", command=lambda: self.create_client_for_import(client_var, client_combo)).pack(side="right")
        
        # Site selection
        site_frame = ttk.LabelFrame(dialog, text="Site", padding="10")
        site_frame.pack(fill="x", padx=20, pady=5)
        
        site_var = tk.StringVar()
        site_combo = ttk.Combobox(site_frame, textvariable=site_var, width=40)
        site_combo.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        ttk.Button(site_frame, text="New", command=lambda: self.create_site_for_import(client_var, site_var, site_combo)).pack(side="right")
        
        # Role selection
        role_frame = ttk.LabelFrame(dialog, text="Role", padding="10")
        role_frame.pack(fill="x", padx=20, pady=5)
        
        role_var = tk.StringVar()
        role_combo = ttk.Combobox(role_frame, textvariable=role_var, width=40, values=[
            "Desktop", "Server", "Workstation", "Domain Controller", "Database", 
            "Web Server", "File Server", "Terminal Server", "Unknown"
        ])
        role_combo.pack(fill="x")
        role_combo.set("Desktop")
        
        # Populate existing data
        clients = self.db.get_clients()
        client_names = [name for _, name, _, _ in clients]
        client_combo['values'] = client_names
        
        def on_client_change(event=None):
            client_name = client_var.get()
            if client_name:
                client_id = None
                for cid, name, _, _ in clients:
                    if name == client_name:
                        client_id = cid
                        break
                if client_id:
                    sites = self.db.get_sites(client_id)
                    site_names = [name for _, _, name, _, _, _ in sites]
                    site_combo['values'] = site_names
        
        client_combo.bind('<<ComboboxSelected>>', on_client_change)
        
        # Buttons
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=20)
        
        def save_import():
            if not client_var.get() or not site_var.get() or not role_var.get():
                messagebox.showerror("Error", "Please fill all fields")
                return
            
            try:
                # Find client and site IDs
                client_id = None
                for cid, name, _, _ in clients:
                    if name == client_var.get():
                        client_id = cid
                        break
                
                if not client_id:
                    messagebox.showerror("Error", "Client not found")
                    return
                
                sites = self.db.get_sites(client_id)
                site_id = None
                for sid, _, name, _, _, _ in sites:
                    if name == site_var.get():
                        site_id = sid
                        break
                
                if not site_id:
                    messagebox.showerror("Error", "Site not found")
                    return
                
                # Generate new UUID and rename file
                new_uuid = generate_uuidv7()
                new_filename = f"{new_uuid}.wim"
                new_path = self.image_store_path / new_filename
                
                # Rename file to UUID format
                wim_file.rename(new_path)
                
                # Get file size
                image_size_gb = int(new_path.stat().st_size / (1024**3))
                
                # Import to database
                image_id = self.db.add_image(
                    client_id, site_id, role_var.get(), 
                    '', str(new_path), image_size_gb
                )
                
                # Save metadata
                self.db.save_image_metadata(image_id, self.image_store_path)
                
                self.log(f"SUCCESS: Imported {wim_file.name} as {new_filename}")
                dialog.destroy()
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to import: {e}")
        
        def skip_import():
            dialog.destroy()
        
        ttk.Button(button_frame, text="Import", command=save_import).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Skip", command=skip_import).pack(side="left", padx=5)

    def create_client_for_import(self, client_var, client_combo):
        """Create new client during import process"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Create New Client")
        dialog.geometry("400x200")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="Client Name:").pack(pady=5)
        name_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=name_var, width=40).pack(pady=5)
        
        ttk.Label(dialog, text="Short Name:").pack(pady=5)
        short_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=short_var, width=40).pack(pady=5)
        
        def save_client():
            name = name_var.get().strip()
            short = short_var.get().strip()
            
            if not name or not short:
                messagebox.showerror("Error", "Both fields are required")
                return
            
            try:
                self.db.add_client(name, short)
                # Refresh client combo
                clients = self.db.get_clients()
                client_names = [n for _, n, _, _ in clients]
                client_combo['values'] = client_names
                client_var.set(name)
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to create client: {e}")
        
        ttk.Button(dialog, text="Create", command=save_client).pack(pady=20)

    def create_site_for_import(self, client_var, site_var, site_combo):
        """Create new site during import process"""
        if not client_var.get():
            messagebox.showerror("Error", "Please select a client first")
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Create New Site")
        dialog.geometry("400x200")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="Site Name:").pack(pady=5)
        name_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=name_var, width=40).pack(pady=5)
        
        ttk.Label(dialog, text="Short Name (for VM naming):").pack(pady=5)
        short_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=short_var, width=40).pack(pady=5)
        
        ttk.Label(dialog, text="Description (optional):").pack(pady=5)
        desc_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=desc_var, width=40).pack(pady=5)
        
        def save_site():
            name = name_var.get().strip()
            short = short_var.get().strip()
            desc = desc_var.get().strip()
            
            if not name or not short:
                messagebox.showerror("Error", "Name and Short Name are required")
                return
            
            try:
                # Find client ID
                clients = self.db.get_clients()
                client_id = None
                for cid, n, _, _ in clients:
                    if n == self.client_var.get():
                        client_id = cid
                        break
                
                if client_id:
                    self.db.add_site(client_id, name, short, desc)
                    # Refresh the parent dialog's site combo
                    self.refresh_client_site_data()
                    # Auto-select the newly created site
                    self.site_var.set(name)
                    # Trigger site selection event
                    self.on_client_selected()
                    self.log(f"SUCCESS: Created new site: {name}")
                    dialog.destroy()
                else:
                    messagebox.showerror("Error", "Could not create site: Client ID not found.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to create site: {e}")
        
        ttk.Button(dialog, text="Create Site", command=save_site).pack(pady=20)

    def open_image_manager(self):
        """Open the Image Manager window"""
        manager_window = tk.Toplevel(self.root)
        manager_window.title("Image Manager")
        manager_window.geometry("1000x700")
        manager_window.transient(self.root)
        
        # Center window
        manager_window.update_idletasks()
        x = (manager_window.winfo_screenwidth() // 2) - (manager_window.winfo_width() // 2)
        y = (manager_window.winfo_screenheight() // 2) - (manager_window.winfo_height() // 2)
        manager_window.geometry(f"+{x}+{y}")
        
        # Header
        header_frame = ttk.Frame(manager_window)
        header_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(header_frame, text="Image Manager", 
                 font=("TkDefaultFont", 16, "bold")).pack(side="left")
        
        ttk.Button(header_frame, text="🔄 Refresh", 
                  command=lambda: self.refresh_image_manager(images_tree)).pack(side="right", padx=(0, 5))
        ttk.Button(header_frame, text="📥 Import WIM", 
                  command=lambda: self.manual_import_wim()).pack(side="right", padx=(0, 5))
        
        # Images tree
        tree_frame = ttk.Frame(manager_window)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        columns = ("Client", "Site", "Role", "Size", "Status", "Created", "File")
        images_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=15)
        
        for col in columns:
            images_tree.heading(col, text=col)
            if col == "File":
                images_tree.column(col, width=200)
            elif col == "Size":
                images_tree.column(col, width=80)
            else:
                images_tree.column(col, width=120)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=images_tree.yview)
        h_scrollbar = ttk.Scrollbar(tree_frame, orient="horizontal", command=images_tree.xview)
        images_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        images_tree.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        
        # Details panel
        details_frame = ttk.LabelFrame(manager_window, text="Image Details", padding="10")
        details_frame.pack(fill="x", padx=10, pady=(0, 10))
        
        self.details_text = tk.Text(details_frame, height=8, wrap=tk.WORD, state='disabled')
        self.details_text.pack(fill="both", expand=True)
        
        # Bind selection event
        def on_image_select(event):
            selection = images_tree.selection()
            if selection:
                item = images_tree.item(selection[0])
                values = item['values']
                if values:
                    self.show_image_details(values, item['tags'][0] if item['tags'] else None)
        
        images_tree.bind('<<TreeviewSelect>>', on_image_select)
        
        # Load initial data
        self.refresh_image_manager(images_tree)

    def refresh_image_manager(self, images_tree):
        """Refresh the image manager tree view"""
        try:
            images_tree.delete(*images_tree.get_children())
            images = self.db.get_images()
            
            for image_data in images:
                (image_id, role, image_path, image_size_gb, vm_name, vm_created, 
                 status, created_at, client_name, client_short, site_name, site_short) = image_data
                
                # Format display data
                created_date = created_at.split()[0] if created_at else "Unknown"
                filename = Path(image_path).name if image_path else "Missing"
                
                images_tree.insert("", "end", values=(
                    client_name, site_name, role, f"{image_size_gb} GB", 
                    status, created_date, filename
                ), tags=(image_id,))
                
        except Exception as e:
            self.log(f"ERROR: Failed to refresh image manager: {e}")

    def show_image_details(self, values, image_id):
        """Show detailed information about selected image"""
        try:
            self.details_text.config(state='normal')
            self.details_text.delete(1.0, tk.END)
            
            if not image_id:
                self.details_text.insert(tk.END, "No image selected")
                self.details_text.config(state='disabled')
                return
            
            # Get full image details from database
            with sqlite3.connect(self.db.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute('''
                    SELECT i.*, c.name as client_name, c.short_name as client_short,
                           s.name as site_name, s.short_name as site_short
                    FROM images i
                    JOIN clients c ON i.client_id = c.id
                    JOIN sites s ON i.site_id = s.id
                    WHERE i.id = ?
                ''', (image_id,))
                
                row = cursor.fetchone()
                if row:
                    columns = [description[0] for description in cursor.description]
                    image_data = dict(zip(columns, row))
                    
                    details = f"""Image Details:

ID: {image_data['id']}
Client: {image_data['client_name']} ({image_data['client_short']})
Site: {image_data['site_name']} ({image_data['site_short']})
Role: {image_data['role']}
Status: {image_data['status']}

File Information:
Path: {image_data['image_path']}
Size: {image_data['image_size_gb']} GB
Source WIM: {image_data['wim_source_path'] or 'N/A'}

Virtual Machine:
VM Name: {image_data['vm_name'] or 'N/A'}
VM Created: {'Yes' if image_data['vm_created'] else 'No'}

Timestamps:
Created: {image_data['created_at']}
Updated: {image_data['updated_at']}"""
                    
                    # Check if file exists
                    if image_data['image_path']:
                        image_path = Path(image_data['image_path'])
                        if image_path.exists():
                            actual_size = image_path.stat().st_size / (1024**3)
                            details += f"\n\nFile Status: ✅ File exists ({actual_size:.1f} GB)"
                        else:
                            details += f"\n\nFile Status: ❌ File missing"
                    
                    self.details_text.insert(tk.END, details)
                else:
                    self.details_text.insert(tk.END, "Image not found in database")
            
            self.details_text.config(state='disabled')
            
        except Exception as e:
            self.details_text.config(state='normal')
            self.details_text.delete(1.0, tk.END)
            self.details_text.insert(tk.END, f"Error loading details: {e}")
            self.details_text.config(state='disabled')

    def manual_import_wim(self):
        """Manual WIM import from file dialog"""
        wim_file = filedialog.askopenfilename(
            title="Select WIM File to Import",
            filetypes=[("WIM Files", "*.wim"), ("All Files", "*.*")]
        )
        
        if wim_file:
            wim_path = Path(wim_file)
            self.import_single_wim_file(wim_path)

    def show_first_time_setup(self):
        """Show first-time setup dialog for initial configuration"""
        setup_window = tk.Toplevel(self.root)
        setup_window.title("First Time Setup - PYC Image Manager")
        setup_window.geometry("600x500")
        setup_window.transient(self.root)
        setup_window.grab_set()
        
        # Center the window
        setup_window.update_idletasks()
        x = (setup_window.winfo_screenwidth() // 2) - (setup_window.winfo_width() // 2)
        y = (setup_window.winfo_screenheight() // 2) - (setup_window.winfo_height() // 2)
        setup_window.geometry(f"+{x}+{y}")
        
        ttk.Label(setup_window, text="Welcome to PYC Image Manager", 
                 font=("TkDefaultFont", 14, "bold")).pack(pady=20)
        ttk.Label(setup_window, text="Let's configure your storage locations", 
                 font=("TkDefaultFont", 10)).pack(pady=5)
        
        # VHDX Storage Path
        path_frame = ttk.LabelFrame(setup_window, text="VHDX Storage Location", padding="10")
        path_frame.pack(fill="x", padx=20, pady=10)
        
        ttk.Label(path_frame, text="Where should we store your VHDX image files?").pack(anchor="w")
        
        path_var = tk.StringVar()
        path_entry = ttk.Entry(path_frame, textvariable=path_var, width=60)
        path_entry.pack(fill="x", pady=5)
        
        def browse_path():
            path = filedialog.askdirectory(title="Select VHDX Storage Directory")
            if path:
                path_var.set(path)
        
        ttk.Button(path_frame, text="Browse...", command=browse_path).pack(pady=5)
        
        # Auto-detect largest drive
        largest_drive = self.find_largest_drive()
        default_path = str(Path(largest_drive) / "pyc-images-dev")
        path_var.set(default_path)
        
        ttk.Label(path_frame, text=f"Default: Largest available drive ({largest_drive})", 
                 font=("TkDefaultFont", 8), foreground="gray").pack(anchor="w")
        
        # Restic Repository Storage Path
        restic_frame = ttk.LabelFrame(setup_window, text="Restic Repository Location", padding="10")
        restic_frame.pack(fill="x", padx=20, pady=10)
        
        ttk.Label(restic_frame, text="Where should we store your restic backup repositories?").pack(anchor="w")
        
        restic_var = tk.StringVar()
        restic_entry = ttk.Entry(restic_frame, textvariable=restic_var, width=60)
        restic_entry.pack(fill="x", pady=5)
        
        def browse_restic_path():
            path = filedialog.askdirectory(title="Select Restic Repository Directory")
            if path:
                restic_var.set(path)
        
        ttk.Button(restic_frame, text="Browse...", command=browse_restic_path).pack(pady=5)
        
        # Auto-detect largest drive for restic
        default_restic_path = str(Path(largest_drive) / "pyc-restic-repo")
        restic_var.set(default_restic_path)
        
        ttk.Label(restic_frame, text=f"Default: {default_restic_path}", 
                 font=("TkDefaultFont", 8), foreground="gray").pack(anchor="w")
        ttk.Label(restic_frame, text="Repositories will be organized by client UUID under this path", 
                 font=("TkDefaultFont", 8), foreground="gray").pack(anchor="w")
        
        # Buttons
        button_frame = ttk.Frame(setup_window)
        button_frame.pack(pady=20)
        
        def save_config():
            vhdx_path = path_var.get().strip()
            restic_path = restic_var.get().strip()
            
            if not vhdx_path:
                messagebox.showerror("Error", "Please select a VHDX storage path")
                return
            
            if not restic_path:
                messagebox.showerror("Error", "Please select a restic repository path")
                return
            
            # Create directories if they don't exist
            try:
                Path(vhdx_path).mkdir(parents=True, exist_ok=True)
                Path(restic_path).mkdir(parents=True, exist_ok=True)
                
                self.db.set_config("image_store_path", vhdx_path)
                self.db.set_config("restic_repository_base_path", restic_path)
                self.db.set_config("first_time_setup_complete", "true")
                
                self.log(f"INFO: VHDX storage configured at: {vhdx_path}")
                self.log(f"INFO: Restic repository base path configured at: {restic_path}")
                setup_window.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to create directories: {e}")
        
        ttk.Button(button_frame, text="Save Configuration", command=save_config).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Cancel", command=setup_window.destroy).pack(side="left", padx=5)
        
        setup_window.wait_window()

    def get_image_store_path(self):
        """Get the image storage path from config"""
        stored_path = self.db.get_config("image_store_path")
        if stored_path and Path(stored_path).exists():
            return Path(stored_path)
        
        # Fallback to default
        largest_drive = self.find_largest_drive()
        default_path = Path(largest_drive) / "pyc-images-dev"
        default_path.mkdir(parents=True, exist_ok=True)
        self.db.set_config("image_store_path", str(default_path))
        return default_path

    def find_largest_drive(self):
        """Find the drive with the most free space"""
        max_free = 0
        largest_drive = "C:\\"
        
        for drive in string.ascii_uppercase:
            drive_path = f"{drive}:\\"
            if Path(drive_path).exists():
                try:
                    free_space = shutil.disk_usage(drive_path).free
                    if free_space > max_free:
                        max_free = free_space
                        largest_drive = drive_path
                except:
                    continue
        
        return largest_drive

    def get_restic_base_path(self):
        """Get the restic repository base path from config"""
        stored_path = self.db.get_config("restic_repository_base_path")
        if stored_path and Path(stored_path).exists():
            return Path(stored_path)
        
        # Fallback to default
        largest_drive = self.find_largest_drive()
        default_path = Path(largest_drive) / "pyc-restic-repo"
        default_path.mkdir(parents=True, exist_ok=True)
        self.db.set_config("restic_repository_base_path", str(default_path))
        return default_path
    
    def check_s3_configuration(self):
        """Check and configure S3 settings if not already set"""
        s3_config = self.db.get_s3_config()
        if not s3_config:
            self.show_s3_configuration_dialog()
    
    def show_s3_configuration_dialog(self):
        """Show S3 configuration dialog"""
        dialog = tk.Toplevel(self.root)
        dialog.title("S3 Configuration Required")
        dialog.geometry("500x400")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # Main frame
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="S3 Cloud Storage Configuration", 
                               font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 10))
        
        # Description
        desc_label = ttk.Label(main_frame, 
                              text="Configure S3 cloud storage for restic repositories.\nThis enables cloud backup and synchronization.",
                              justify=tk.CENTER)
        desc_label.pack(pady=(0, 20))
        
        # Configuration fields
        fields_frame = ttk.Frame(main_frame)
        fields_frame.pack(fill=tk.BOTH, expand=True)
        
        # S3 Bucket
        ttk.Label(fields_frame, text="S3 Bucket Name:").grid(row=0, column=0, sticky="w", pady=5)
        bucket_var = tk.StringVar()
        bucket_entry = ttk.Entry(fields_frame, textvariable=bucket_var, width=40)
        bucket_entry.grid(row=0, column=1, sticky="ew", pady=5, padx=(10, 0))
        
        # Access Key
        ttk.Label(fields_frame, text="Access Key ID:").grid(row=1, column=0, sticky="w", pady=5)
        access_key_var = tk.StringVar()
        access_key_entry = ttk.Entry(fields_frame, textvariable=access_key_var, width=40)
        access_key_entry.grid(row=1, column=1, sticky="ew", pady=5, padx=(10, 0))
        
        # Secret Key
        ttk.Label(fields_frame, text="Secret Access Key:").grid(row=2, column=0, sticky="w", pady=5)
        secret_key_var = tk.StringVar()
        secret_key_entry = ttk.Entry(fields_frame, textvariable=secret_key_var, width=40, show="*")
        secret_key_entry.grid(row=2, column=1, sticky="ew", pady=5, padx=(10, 0))
        
        # Endpoint
        ttk.Label(fields_frame, text="S3 Endpoint:").grid(row=3, column=0, sticky="w", pady=5)
        endpoint_var = tk.StringVar(value="s3.amazonaws.com")
        endpoint_entry = ttk.Entry(fields_frame, textvariable=endpoint_var, width=40)
        endpoint_entry.grid(row=3, column=1, sticky="ew", pady=5, padx=(10, 0))
        
        # Region
        ttk.Label(fields_frame, text="S3 Region:").grid(row=4, column=0, sticky="w", pady=5)
        region_var = tk.StringVar(value="us-east-1")
        region_entry = ttk.Entry(fields_frame, textvariable=region_var, width=40)
        region_entry.grid(row=4, column=1, sticky="ew", pady=5, padx=(10, 0))
        
        # Configure grid weights
        fields_frame.columnconfigure(1, weight=1)
        
        # Buttons frame
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, pady=(20, 0))
        
        def save_s3_config():
            # Validate required fields
            if not all([bucket_var.get().strip(), access_key_var.get().strip(), 
                       secret_key_var.get().strip(), endpoint_var.get().strip()]):
                messagebox.showerror("Error", "All fields are required")
                return
            
            try:
                # Save S3 configuration
                self.db.set_s3_config(
                    bucket=bucket_var.get().strip(),
                    access_key=access_key_var.get().strip(),
                    secret_key=secret_key_var.get().strip(),
                    endpoint=endpoint_var.get().strip(),
                    region=region_var.get().strip() or "us-east-1"
                )
                messagebox.showinfo("Success", "S3 configuration saved successfully!")
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save S3 configuration: {str(e)}")
        
        def skip_s3_config():
            result = messagebox.askyesno("Skip S3 Configuration", 
                                       "Are you sure you want to skip S3 configuration?\n\n" +
                                       "You can configure S3 later, but cloud storage features will not be available.")
            if result:
                dialog.destroy()
        
        # Buttons
        ttk.Button(buttons_frame, text="Save Configuration", command=save_s3_config).pack(side=tk.RIGHT, padx=(10, 0))
        ttk.Button(buttons_frame, text="Skip for Now", command=skip_s3_config).pack(side=tk.RIGHT)
    
    def check_workflow_mode(self):
        """Check and configure workflow mode (development vs production)"""
        workflow_mode = self.db.get_config('workflow_mode')
        if not workflow_mode:
            self.show_workflow_mode_dialog()
    
    def show_workflow_mode_dialog(self):
        """Show workflow mode selection dialog"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Workflow Mode Selection")
        dialog.geometry("500x350")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # Main frame
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="Select Your Workflow Mode", 
                               font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 10))
        
        # Description
        desc_label = ttk.Label(main_frame, 
                              text="How are you using this imaging tool?",
                              justify=tk.CENTER)
        desc_label.pack(pady=(0, 20))
        
        # Mode selection
        mode_var = tk.StringVar(value="development")
        
        # Development mode option
        dev_frame = ttk.LabelFrame(main_frame, text="Development Mode", padding="10")
        dev_frame.pack(fill="x", pady=5)
        
        ttk.Radiobutton(dev_frame, text="I'm developing/creating a new image", 
                       variable=mode_var, value="development").pack(anchor="w")
        
        dev_desc = ttk.Label(dev_frame, 
                            text="• Tool is being used temporarily on this machine\n" +
                                 "• Creating images for later deployment\n" +
                                 "• All snapshots will be tagged as 'development'\n" +
                                 "• Repositories can be imported to production systems later",
                            font=("TkDefaultFont", 9), justify="left")
        dev_desc.pack(anchor="w", pady=(5, 0))
        
        # Production mode option
        prod_frame = ttk.LabelFrame(main_frame, text="Production Mode", padding="10")
        prod_frame.pack(fill="x", pady=5)
        
        ttk.Radiobutton(prod_frame, text="I'm managing the image system permanently", 
                       variable=mode_var, value="production").pack(anchor="w")
        
        prod_desc = ttk.Label(prod_frame, 
                             text="• This machine is the permanent image management system\n" +
                                  "• Managing production images and deployments\n" +
                                  "• All snapshots will be tagged as 'production'\n" +
                                  "• Full repository management and restoration capabilities",
                             font=("TkDefaultFont", 9), justify="left")
        prod_desc.pack(anchor="w", pady=(5, 0))
        
        # Buttons frame
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, pady=(20, 0))
        
        def save_workflow_mode():
            selected_mode = mode_var.get()
            try:
                self.db.set_config('workflow_mode', selected_mode)
                messagebox.showinfo("Success", f"Workflow mode set to: {selected_mode.title()}")
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save workflow mode: {str(e)}")
        
        # Buttons
        ttk.Button(buttons_frame, text="OK", command=save_workflow_mode).pack(side=tk.RIGHT)
    
    def detect_workflow_mode(self):
        """Detect workflow mode by checking both potential database locations"""
        # Check production location first
        prod_db_path = Path(os.environ.get('PUBLIC', 'C:\\Users\\Public')) / "Documents" / "pyc" / "pyc.db"
        if prod_db_path.exists():
            try:
                with sqlite3.connect(prod_db_path) as conn:
                    cursor = conn.cursor()
                    cursor.execute('SELECT value FROM config WHERE key = ?', ('workflow_mode',))
                    result = cursor.fetchone()
                    if result:
                        return result[0]
            except:
                pass
        
        # Check development location
        windir = os.environ.get('WINDIR', 'C:\\Windows')
        dev_db_path = Path(windir) / "Temp" / "pyc.db"
        if dev_db_path.exists():
            try:
                with sqlite3.connect(dev_db_path) as conn:
                    cursor = conn.cursor()
                    cursor.execute('SELECT value FROM config WHERE key = ?', ('workflow_mode',))
                    result = cursor.fetchone()
                    if result:
                        return result[0]
            except:
                pass
        
        # Default to development mode if no config found
        return "development"
    
    def get_workflow_mode(self):
        """Get the current workflow mode"""
        return self.db.get_config('workflow_mode', 'development')
    
    def on_image_type_changed(self):
        """Handle image type selection change (new vs existing)"""
        if not hasattr(self, 'image_type_var'):
            return
            
        image_type = self.image_type_var.get()
        
        if image_type == "existing":
            # Show existing image selection
            if hasattr(self, 'existing_image_frame'):
                self.existing_image_frame.grid()
                self.refresh_existing_images()
        else:
            # Hide existing image selection
            if hasattr(self, 'existing_image_frame'):
                self.existing_image_frame.grid_remove()
    
    def refresh_existing_images(self):
        """Refresh the list of existing images for the selected client and role"""
        if not hasattr(self, 'existing_image_combo'):
            return
            
        try:
            # Get selected client and role
            client_name = self.client_var.get() if hasattr(self, 'client_var') else ""
            role = self.role_var.get() if hasattr(self, 'role_var') else ""
            
            if not client_name or client_name == "-- Select Client --":
                self.existing_image_combo['values'] = []
                return
            
            # Find client ID
            clients = self.db.get_clients()
            client_id = None
            for cid, name, _, _ in clients:
                if name == client_name:
                    client_id = cid
                    break
            
            if not client_id:
                self.existing_image_combo['values'] = []
                return
            
            # Get existing images for this client and role
            images = self.db.get_images()
            matching_images = []
            
            for image in images:
                # image structure: (id, client_id, site_id, role, repository_path, ...)
                if len(image) >= 4 and image[1] == client_id and image[3] == role:
                    # Format: "Role - Site - ImageID (snapshots)"
                    image_id = image[0]
                    site_id = image[2]
                    snapshot_count = image[5] if len(image) > 5 else 0
                    
                    # Get site name
                    sites = self.db.get_sites(client_id)
                    site_name = "Unknown Site"
                    for site in sites:
                        if site[0] == site_id:
                            site_name = site[2]  # site name is at index 2
                            break
                    
                    display_text = f"{role} - {site_name} - {image_id[:8]}... ({snapshot_count} snapshots)"
                    matching_images.append((display_text, image_id))
            
            # Update combo box
            if matching_images:
                display_values = [item[0] for item in matching_images]
                self.existing_image_combo['values'] = display_values
                # Store the mapping for retrieval
                if not hasattr(self, 'image_id_mapping'):
                    self.image_id_mapping = {}
                self.image_id_mapping = {item[0]: item[1] for item in matching_images}
            else:
                self.existing_image_combo['values'] = ["No existing images found"]
                self.image_id_mapping = {}
                
        except Exception as e:
            print(f"Error refreshing existing images: {e}")
            self.existing_image_combo['values'] = ["Error loading images"]
    
    def get_selected_image_uuid(self):
        """Get the UUID of the currently selected existing image"""
        if not hasattr(self, 'existing_image_var') or not hasattr(self, 'image_id_mapping'):
            return None
            
        selected_display = self.existing_image_var.get()
        return self.image_id_mapping.get(selected_display)
    
    def show_password_manager_reminder(self, password, identifier, client_name="", site_name="", role=""):
        """Show password manager reminder dialog"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Important: Save Password")
        dialog.geometry("600x400")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # Main frame
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Warning icon and title
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill="x", pady=(0, 15))
        
        ttk.Label(title_frame, text="⚠️", font=("Arial", 24)).pack(side="left", padx=(0, 10))
        ttk.Label(title_frame, text="IMPORTANT: Save This Password", 
                 font=("Arial", 14, "bold"), foreground="red").pack(side="left")
        
        # Instructions
        instructions = ttk.Label(main_frame, 
                               text="Please save this repository password to your password manager immediately.\n" +
                                    "You will need this password to access the backup repository later.",
                               font=("TkDefaultFont", 10), justify="left")
        instructions.pack(anchor="w", pady=(0, 15))
        
        # Client/Site/Role info
        if client_name or site_name or role:
            info_frame = ttk.LabelFrame(main_frame, text="Repository Information", padding="10")
            info_frame.pack(fill="x", pady=(0, 15))
            
            if client_name:
                ttk.Label(info_frame, text=f"Client: {client_name}", font=("TkDefaultFont", 9)).pack(anchor="w")
            if site_name:
                ttk.Label(info_frame, text=f"Site: {site_name}", font=("TkDefaultFont", 9)).pack(anchor="w")
            if role:
                ttk.Label(info_frame, text=f"Role: {role}", font=("TkDefaultFont", 9)).pack(anchor="w")
        
        # Password section
        password_frame = ttk.LabelFrame(main_frame, text="Password to Save", padding="10")
        password_frame.pack(fill="x", pady=(0, 15))
        
        # Identifier
        ttk.Label(password_frame, text="Suggested Password Manager Entry Name:", font=("TkDefaultFont", 9, "bold")).pack(anchor="w")
        identifier_entry = ttk.Entry(password_frame, font=("Consolas", 10))
        identifier_entry.pack(fill="x", pady=(2, 10))
        identifier_entry.insert(0, identifier)
        identifier_entry.config(state="readonly")
        
        # Password
        ttk.Label(password_frame, text="Repository Password:", font=("TkDefaultFont", 9, "bold")).pack(anchor="w")
        password_entry = ttk.Entry(password_frame, font=("Consolas", 10))
        password_entry.pack(fill="x", pady=(2, 0))
        password_entry.insert(0, password)
        password_entry.config(state="readonly")
        
        # Copy buttons
        copy_frame = ttk.Frame(password_frame)
        copy_frame.pack(fill="x", pady=(5, 0))
        
        def copy_identifier():
            dialog.clipboard_clear()
            dialog.clipboard_append(identifier)
            messagebox.showinfo("Copied", "Identifier copied to clipboard")
        
        def copy_password():
            dialog.clipboard_clear()
            dialog.clipboard_append(password)
            messagebox.showinfo("Copied", "Password copied to clipboard")
        
        ttk.Button(copy_frame, text="Copy Identifier", command=copy_identifier).pack(side="left", padx=(0, 10))
        ttk.Button(copy_frame, text="Copy Password", command=copy_password).pack(side="left")
        
        # Warning message
        warning_label = ttk.Label(main_frame,
                                text="⚠️ This password will be stored securely in the database, but please save it to your\n" +
                                     "password manager as a backup. This is especially important for development mode\n" +
                                     "as the database may be cleared during generalization.",
                                font=("TkDefaultFont", 9), foreground="red", justify="left")
        warning_label.pack(anchor="w", pady=(0, 15))
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x")
        
        def acknowledge():
            dialog.destroy()
        
        ttk.Button(button_frame, text="I Have Saved The Password", command=acknowledge).pack(side="right")
    
    def create_client_metadata_json(self, client_uuid, client_info=None, site_info=None, image_info=None):
        """Create or update JSON metadata file in client repository folder"""
        try:
            # Get the client repository path
            restic_base = self.get_restic_base_path()
            client_repo_path = restic_base / client_uuid
            metadata_file = client_repo_path / "client_metadata.json"
            
            # Initialize metadata structure
            metadata = {
                "client": {},
                "sites": {},
                "images": {},
                "last_updated": datetime.now().isoformat()
            }
            
            # Load existing metadata if file exists
            if metadata_file.exists():
                try:
                    with open(metadata_file, 'r', encoding='utf-8') as f:
                        existing_metadata = json.load(f)
                        # Preserve existing data and merge with new
                        metadata.update(existing_metadata)
                except (json.JSONDecodeError, IOError):
                    # If file is corrupted, start fresh
                    pass
            
            # Update client information
            if client_info:
                metadata["client"] = {
                    "id": client_info.get("id", client_uuid),
                    "name": client_info.get("name", ""),
                    "short_name": client_info.get("short_name", ""),
                    "description": client_info.get("description", ""),
                    "created_at": client_info.get("created_at", datetime.now().isoformat()),
                    "updated_at": datetime.now().isoformat()
                }
            
            # Update site information
            if site_info:
                site_id = site_info.get("id")
                if site_id:
                    if "sites" not in metadata:
                        metadata["sites"] = {}
                    metadata["sites"][site_id] = {
                        "id": site_id,
                        "name": site_info.get("name", ""),
                        "short_name": site_info.get("short_name", ""),
                        "description": site_info.get("description", ""),
                        "created_at": site_info.get("created_at", datetime.now().isoformat()),
                        "updated_at": datetime.now().isoformat()
                    }
            
            # Update image information
            if image_info:
                image_id = image_info.get("id")
                if image_id:
                    if "images" not in metadata:
                        metadata["images"] = {}
                    metadata["images"][image_id] = {
                        "id": image_id,
                        "role": image_info.get("role", ""),
                        "site_id": image_info.get("site_id", ""),
                        "repository_path": image_info.get("repository_path", ""),
                        "snapshot_count": image_info.get("snapshot_count", 0),
                        "latest_snapshot_id": image_info.get("latest_snapshot_id", ""),
                        "repository_size_gb": image_info.get("repository_size_gb", 0),
                        "created_at": image_info.get("created_at", datetime.now().isoformat()),
                        "updated_at": datetime.now().isoformat()
                    }
            
            # Update last_updated timestamp
            metadata["last_updated"] = datetime.now().isoformat()
            
            # Ensure the client repository directory exists
            client_repo_path.mkdir(parents=True, exist_ok=True)
            
            # Write the JSON file
            with open(metadata_file, 'w', encoding='utf-8') as f:
                json.dump(metadata, f, indent=2, ensure_ascii=False)
            
            return True
            
        except Exception as e:
            print(f"Failed to create/update client metadata JSON: {str(e)}")
            return False
    
    def load_client_metadata_json(self, client_uuid):
        """Load JSON metadata file from client repository folder"""
        try:
            restic_base = self.get_restic_base_path()
            client_repo_path = restic_base / client_uuid
            metadata_file = client_repo_path / "client_metadata.json"
            
            if metadata_file.exists():
                with open(metadata_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return None
            
        except Exception as e:
            print(f"Failed to load client metadata JSON: {str(e)}")
            return None
    
    def create_s3_image_metadata(self, image_uuid, client_info=None, site_info=None, image_info=None, workflow_mode="development"):
        """Create image-specific metadata JSON file in S3 bucket root metadata folder"""
        try:
            # Check if we're using S3
            repo_type = self.repo_type_var.get() if hasattr(self, 'repo_type_var') else "local"
            if repo_type != "s3":
                return True  # Not using S3, skip
            
            # Get S3 configuration
            s3_config = self.db.get_s3_config()
            if not s3_config:
                return False
            
            # Create complete image metadata - self-contained for import
            metadata = {
                "format_version": "1.0",
                "metadata_type": "image",
                "image": {
                    "id": image_uuid,
                    "role": image_info.get("role", "") if image_info else "",
                    "workflow_mode": workflow_mode,
                    "repository_path": image_info.get("repository_path", "") if image_info else "",
                    "snapshot_count": image_info.get("snapshot_count", 0) if image_info else 0,
                    "latest_snapshot_id": image_info.get("latest_snapshot_id", "") if image_info else "",
                    "repository_size_gb": image_info.get("repository_size_gb", 0) if image_info else 0,
                    "created_at": image_info.get("created_at", datetime.now().isoformat()) if image_info else datetime.now().isoformat(),
                    "updated_at": datetime.now().isoformat()
                },
                "client": {
                    "id": client_info.get("id", "") if client_info else "",
                    "name": client_info.get("name", "") if client_info else "",
                    "short_name": client_info.get("short_name", "") if client_info else "",
                    "description": client_info.get("description", "") if client_info else ""
                },
                "site": {
                    "id": site_info.get("id", "") if site_info else "",
                    "name": site_info.get("name", "") if site_info else "",
                    "short_name": site_info.get("short_name", "") if site_info else "",
                    "description": site_info.get("description", "") if site_info else ""
                },
                "export_info": {
                    "exported_from": os.environ.get('COMPUTERNAME', 'unknown'),
                    "export_timestamp": datetime.now().isoformat(),
                    "tool_version": "1.0"
                }
            }
            
            # Create temporary file for upload
            import tempfile
            with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as temp_file:
                json.dump(metadata, temp_file, indent=2, ensure_ascii=False)
                temp_file_path = temp_file.name
            
            # Construct S3 path for image metadata
            s3_bucket = s3_config.get('s3_bucket')
            s3_endpoint = s3_config.get('s3_endpoint', 's3.amazonaws.com')
            metadata_key = f"metadata/{image_uuid}.json"
            
            # Set environment variables for AWS CLI
            access_key = s3_config.get('s3_access_key')
            secret_key = s3_config.get('s3_secret_key')
            region = s3_config.get('s3_region', 'us-east-1')
            
            if access_key:
                os.environ['AWS_ACCESS_KEY_ID'] = access_key
            if secret_key:
                os.environ['AWS_SECRET_ACCESS_KEY'] = secret_key
            if region:
                os.environ['AWS_DEFAULT_REGION'] = region
            
            # Upload using AWS CLI
            upload_cmd = [
                'aws', 's3', 'cp', 
                temp_file_path, 
                f"s3://{s3_bucket}/{metadata_key}",
                '--region', s3_config.get('s3_region', 'us-east-1')
            ]
            
            # If not using AWS, add endpoint
            if s3_endpoint != 's3.amazonaws.com':
                upload_cmd.extend(['--endpoint-url', f"https://{s3_endpoint}"])
            
            result = subprocess.run(upload_cmd, capture_output=True, text=True)
            
            # Clean up temp file
            os.unlink(temp_file_path)
            
            if result.returncode == 0:
                print(f"Successfully uploaded image metadata to S3: {metadata_key}")
                return True
            else:
                print(f"Failed to upload image metadata to S3: {result.stderr}")
                return False
                
        except Exception as e:
            print(f"Failed to create S3 image metadata: {str(e)}")
            return False

    def create_workflow_header(self):
        """Creates the workflow progress header."""
        header_frame = ttk.LabelFrame(self.root, text="Imaging Workflow Progress", padding="10")
        header_frame.pack(fill="x", padx=10, pady=10)
        
        # Step indicators
        steps_frame = ttk.Frame(header_frame)
        steps_frame.pack(fill="x")
        
        self.step_labels = {}
        self.step_buttons = {}
        step_names = [
            "Create System Backup",
            "Professional Image & VM Management", 
            "Generalize & Cleanup"
        ]
        
        for i, step_name in enumerate(step_names, 1):
            # Step frame with clickable button
            step_frame = ttk.Frame(steps_frame)
            step_frame.pack(side="left", fill="x", expand=True, padx=2)
            
            # Clickable step button
            step_text = f"{i}. {step_name}"
            self.step_buttons[i] = ttk.Button(step_frame, text=step_text, 
                                            command=lambda step=i: self.show_step(step),
                                            width=20)
            self.step_buttons[i].pack(pady=2)
            
            # Store reference for styling updates
            self.step_labels[i] = self.step_buttons[i]
        
        # Current step indicator
        self.current_step_label = ttk.Label(header_frame, text="Current Step: 1 - Create System Backup", 
                                          font=("TkDefaultFont", 11, "bold"), foreground="blue")
        self.current_step_label.pack(pady=(10, 0))
        
        # Navigation instructions
        nav_help = ttk.Label(header_frame, text="💡 Click any step above to jump directly to it, or use the navigation buttons below", 
                           font=("TkDefaultFont", 8), foreground="gray")
        nav_help.pack(pady=(5, 0))

    def create_navigation_controls(self):
        """Creates the navigation buttons."""
        nav_frame = ttk.LabelFrame(self.root, text="Navigation Controls", padding="8")
        nav_frame.pack(fill="x", padx=10, pady=5)
        
        # Left side - Previous button
        left_frame = ttk.Frame(nav_frame)
        left_frame.pack(side="left")
        
        self.prev_button = ttk.Button(left_frame, text="← Previous Step", command=self.previous_step, width=15)
        self.prev_button.pack(side="left")
        
        # Add keyboard shortcut labels
        ttk.Label(left_frame, text="(Ctrl+Left)", font=("TkDefaultFont", 7), foreground="gray").pack(side="left", padx=(5, 0))
        
        # Right side - Next button  
        right_frame = ttk.Frame(nav_frame)
        right_frame.pack(side="right")
        
        ttk.Label(right_frame, text="(Ctrl+Right)", font=("TkDefaultFont", 7), foreground="gray").pack(side="right", padx=(0, 5))
        self.next_button = ttk.Button(right_frame, text="Next Step →", command=self.next_step, width=15)
        self.next_button.pack(side="right")
        
        # Center - Quick jump buttons
        center_frame = ttk.Frame(nav_frame)
        center_frame.pack()
        
        ttk.Label(center_frame, text="Quick Jump:", font=("TkDefaultFont", 9, "bold")).pack(side="left", padx=(0, 8))
        
        for i in range(1, self.total_steps + 1):
            btn = ttk.Button(center_frame, text=f"Step {i}", width=8, 
                           command=lambda step=i: self.show_step(step))
            btn.pack(side="left", padx=2)
            
        # Add keyboard shortcut info
        shortcut_info = ttk.Label(center_frame, text="(Ctrl+1-5)", font=("TkDefaultFont", 7), foreground="gray")
        shortcut_info.pack(side="left", padx=(10, 0))

    def setup_keyboard_shortcuts(self):
        """Sets up keyboard shortcuts for navigation."""
        # Bind keyboard shortcuts
        self.root.bind('<Control-Left>', lambda e: self.previous_step())
        self.root.bind('<Control-Right>', lambda e: self.next_step())
        
        # Bind number keys for direct step access
        for i in range(1, self.total_steps + 1):
            self.root.bind(f'<Control-Key-{i}>', lambda e, step=i: self.show_step(step))
        
        # Make sure the root window can receive focus for keyboard events
        self.root.focus_set()

    def create_mode_selection_screen(self):
        """Creates the main mode selection screen with 4 mode buttons"""
        # Create main mode selection frame
        self.mode_selection_frame = ttk.Frame(self.root)
        self.mode_selection_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Title
        title_label = ttk.Label(self.mode_selection_frame, text="Windows Image Preparation Tool", 
                               font=("TkDefaultFont", 18, "bold"))
        title_label.pack(pady=(0, 30))
        
        # Subtitle
        subtitle_label = ttk.Label(self.mode_selection_frame, text="Select Operating Mode", 
                                  font=("TkDefaultFont", 12))
        subtitle_label.pack(pady=(0, 40))
        
        # Create grid for mode buttons
        button_frame = ttk.Frame(self.mode_selection_frame)
        button_frame.pack(expand=True)
        
        # Mode buttons (2x2 grid)
        develop_btn = ttk.Button(button_frame, text="🔧 DEVELOP CAPTURE", 
                                command=lambda: self.enter_mode("develop_capture"),
                                width=25, style="Large.TButton")
        develop_btn.grid(row=0, column=0, padx=15, pady=15, sticky="nsew")
        
        production_btn = ttk.Button(button_frame, text="🚀 PRODUCTION CAPTURE", 
                                   command=lambda: self.enter_mode("production_capture"),
                                   width=25, style="Large.TButton")
        production_btn.grid(row=0, column=1, padx=15, pady=15, sticky="nsew")
        
        generalize_btn = ttk.Button(button_frame, text="🛠️ GENERALIZE", 
                                   command=lambda: self.enter_mode("generalize"),
                                   width=25, style="Large.TButton")
        generalize_btn.grid(row=1, column=0, padx=15, pady=15, sticky="nsew")
        
        manage_btn = ttk.Button(button_frame, text="📁 MANAGE IMAGES", 
                               command=lambda: self.enter_mode("manage_images"),
                               width=25, style="Large.TButton")
        manage_btn.grid(row=1, column=1, padx=15, pady=15, sticky="nsew")
        
        # Configure grid weights for equal sizing
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)
        button_frame.grid_rowconfigure(0, weight=1)
        button_frame.grid_rowconfigure(1, weight=1)
        
        # Configure large button style
        self.style.configure("Large.TButton", font=("TkDefaultFont", 14, "bold"), padding=20)
        
        # Mode descriptions
        desc_frame = ttk.Frame(self.mode_selection_frame)
        desc_frame.pack(fill="x", pady=(30, 0))
        
        descriptions = [
            "DEVELOP CAPTURE: Create development images with S3 storage integration",
            "PRODUCTION CAPTURE: Create production images for deployment", 
            "GENERALIZE: Prepare images for deployment and cleanup",
            "MANAGE IMAGES: Browse, import, and manage existing images"
        ]
        
        for desc in descriptions:
            ttk.Label(desc_frame, text=desc, font=("TkDefaultFont", 9), foreground="gray").pack(anchor="w", pady=2)

    def enter_mode(self, mode):
        """Enter the specified mode and create its UI"""
        self.current_mode = mode
        
        # Set workflow mode and reinitialize database if needed
        if mode == "develop_capture":
            # Reinitialize database manager for development mode
            self.db_manager = DatabaseManager("development")
            self.db = self.db_manager  # Keep backward compatibility
            self.log("INFO: Switched to development mode - using temp database")
        elif mode == "production_capture":
            # Reinitialize database manager for production mode
            self.db_manager = DatabaseManager("production")
            self.db = self.db_manager  # Keep backward compatibility
            self.log("INFO: Switched to production mode - using permanent database")
        
        # Hide mode selection screen
        self.mode_selection_frame.pack_forget()
        
        # Create back button
        self.create_back_to_modes_button()
        
        # Create mode-specific UI
        if mode == "develop_capture":
            self.create_develop_capture_ui()
        elif mode == "production_capture":
            self.create_production_capture_ui()
        elif mode == "generalize":
            self.create_generalize_ui()
        elif mode == "manage_images":
            self.create_manage_images_ui()
    
    def create_back_to_modes_button(self):
        """Create a back button to return to mode selection"""
        self.back_frame = ttk.Frame(self.root)
        self.back_frame.pack(fill="x", padx=10, pady=5)
        
        back_btn = ttk.Button(self.back_frame, text="← Back to Mode Selection", 
                             command=self.return_to_mode_selection, width=20)
        back_btn.pack(side="left")
        
        # Show current mode
        mode_label = ttk.Label(self.back_frame, text=f"Mode: {self.current_mode.replace('_', ' ').title()}", 
                              font=("TkDefaultFont", 10, "bold"))
        mode_label.pack(side="right")
    
    def return_to_mode_selection(self):
        """Return to the main mode selection screen"""
        # Hide current mode UI
        if hasattr(self, 'back_frame'):
            self.back_frame.pack_forget()
        
        # Hide any mode frames
        for frame in self.mode_frames.values():
            frame.pack_forget()
        
        # Show mode selection screen
        self.mode_selection_frame.pack(fill="both", expand=True, padx=20, pady=20)
        self.current_mode = None

    def create_develop_capture_ui(self):
        """Create the DEVELOP CAPTURE mode UI - S3-dependent development image capture"""
        # Set repository type for development mode (always S3)
        self.repo_type_var = tk.StringVar(value="s3")
        
        # Create main frame for develop capture mode
        develop_frame = ttk.Frame(self.root)
        develop_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self.mode_frames["develop_capture"] = develop_frame
        
        # S3 Configuration Section (mandatory for development mode)
        s3_frame = ttk.LabelFrame(develop_frame, text="S3 Configuration (Required)", padding="10")
        s3_frame.pack(fill="x", pady=(0, 10))
        
        # S3 settings grid
        ttk.Label(s3_frame, text="S3 Bucket:").grid(row=0, column=0, sticky="w", pady=2)
        self.dev_s3_bucket_var = tk.StringVar()
        ttk.Entry(s3_frame, textvariable=self.dev_s3_bucket_var, width=30).grid(row=0, column=1, sticky="w", padx=5)
        
        ttk.Label(s3_frame, text="Access Key:").grid(row=1, column=0, sticky="w", pady=2)
        self.dev_s3_access_var = tk.StringVar()
        ttk.Entry(s3_frame, textvariable=self.dev_s3_access_var, width=30).grid(row=1, column=1, sticky="w", padx=5)
        
        ttk.Label(s3_frame, text="Secret Key:").grid(row=2, column=0, sticky="w", pady=2)
        self.dev_s3_secret_var = tk.StringVar()
        ttk.Entry(s3_frame, textvariable=self.dev_s3_secret_var, width=30, show="*").grid(row=2, column=1, sticky="w", padx=5)
        
        ttk.Label(s3_frame, text="S3 Endpoint:").grid(row=3, column=0, sticky="w", pady=2)
        self.dev_s3_endpoint_var = tk.StringVar(value="s3.amazonaws.com")
        ttk.Entry(s3_frame, textvariable=self.dev_s3_endpoint_var, width=30).grid(row=3, column=1, sticky="w", padx=5)
        
        ttk.Label(s3_frame, text="Region:").grid(row=4, column=0, sticky="w", pady=2)
        self.dev_s3_region_var = tk.StringVar(value="us-east-1")
        ttk.Entry(s3_frame, textvariable=self.dev_s3_region_var, width=30).grid(row=4, column=1, sticky="w", padx=5)
        
        # Load and scan S3 button
        ttk.Button(s3_frame, text="Load S3 Metadata & Scan Images", 
                  command=self.load_s3_and_scan_dev_mode).grid(row=5, column=0, columnspan=2, pady=10)
        
        # Client/Site/Image Selection Section
        selection_frame = ttk.LabelFrame(develop_frame, text="Image Configuration", padding="10")
        selection_frame.pack(fill="x", pady=(0, 10))
        
        # Client selection/creation
        ttk.Label(selection_frame, text="Client:").grid(row=0, column=0, sticky="w", pady=2)
        self.dev_client_var = tk.StringVar()
        self.dev_client_combo = ttk.Combobox(selection_frame, textvariable=self.dev_client_var, width=25)
        self.dev_client_combo.grid(row=0, column=1, sticky="w", padx=5)
        self.dev_client_combo.bind('<<ComboboxSelected>>', self.on_dev_client_selected)
        
        ttk.Button(selection_frame, text="New Client", 
                  command=self.create_new_dev_client, width=12).grid(row=0, column=2, padx=5)
        
        # Site selection/creation  
        ttk.Label(selection_frame, text="Site:").grid(row=1, column=0, sticky="w", pady=2)
        self.dev_site_var = tk.StringVar()
        self.dev_site_combo = ttk.Combobox(selection_frame, textvariable=self.dev_site_var, width=25)
        self.dev_site_combo.grid(row=1, column=1, sticky="w", padx=5)
        
        ttk.Button(selection_frame, text="New Site", 
                  command=self.create_new_dev_site, width=12).grid(row=1, column=2, padx=5)
        
        # Role selection
        ttk.Label(selection_frame, text="Role:").grid(row=2, column=0, sticky="w", pady=2)
        self.dev_role_var = tk.StringVar()
        role_combo = ttk.Combobox(selection_frame, textvariable=self.dev_role_var, width=25,
                                 values=["ADMIN", "OP", "MANAGER", "VIP", "KIOSK", "SERVER", "IMAGING"])
        role_combo.grid(row=2, column=1, sticky="w", padx=5)
        
        # Existing images for this client (auto-populated from S3)
        images_frame = ttk.LabelFrame(develop_frame, text="Existing Development Images", padding="10")
        images_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        # Images listbox with scrollbar
        list_frame = ttk.Frame(images_frame)
        list_frame.pack(fill="both", expand=True)
        
        self.dev_images_listbox = tk.Listbox(list_frame, height=8)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.dev_images_listbox.yview)
        self.dev_images_listbox.configure(yscrollcommand=scrollbar.set)
        
        self.dev_images_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        self.dev_images_listbox.bind('<<ListboxSelect>>', self.on_dev_image_selected)
        
        # Action buttons
        action_frame = ttk.Frame(develop_frame)
        action_frame.pack(fill="x", pady=10)
        
        # Create new development image button
        ttk.Button(action_frame, text="Create New Development Image", 
                  command=self.create_dev_image, 
                  style="Large.TButton").pack(side="left", padx=(0, 10))
        
        # Update existing image button  
        ttk.Button(action_frame, text="Update Selected Image", 
                  command=self.update_dev_image,
                  style="Large.TButton").pack(side="left")
        
        # Load S3 configuration if available
        self.load_dev_s3_config()

    def load_dev_s3_config(self):
        """Load S3 configuration for development mode"""
        try:
            s3_config = self.db_manager.get_config("s3_config")
            if s3_config:
                config = json.loads(s3_config)
                self.dev_s3_bucket_var.set(config.get("s3_bucket", ""))
                self.dev_s3_access_var.set(config.get("s3_access_key", ""))
                self.dev_s3_secret_var.set(config.get("s3_secret_key", ""))
                self.dev_s3_endpoint_var.set(config.get("s3_endpoint", "s3.amazonaws.com"))
                self.dev_s3_region_var.set(config.get("s3_region", "us-east-1"))
                self.log(f"INFO: Loaded existing S3 configuration from development database")
                
                # Auto-scan S3 if configuration exists
                if all([config.get("s3_bucket"), config.get("s3_access_key"), config.get("s3_secret_key")]):
                    self.log("INFO: Auto-scanning S3 for existing development images...")
                    threading.Thread(target=self.scan_s3_for_dev_images, daemon=True).start()
            else:
                self.log("INFO: No existing S3 configuration found in development database")
        except Exception as e:
            self.log(f"INFO: No existing S3 configuration found: {e}")

    def load_s3_and_scan_dev_mode(self):
        """Load S3 configuration and scan for existing development images"""
        try:
            # Save S3 configuration
            s3_config = {
                "s3_bucket": self.dev_s3_bucket_var.get(),
                "s3_access_key": self.dev_s3_access_var.get(),
                "s3_secret_key": self.dev_s3_secret_var.get(),
                "s3_endpoint": self.dev_s3_endpoint_var.get(),
                "s3_region": self.dev_s3_region_var.get()
            }
            
            # Validate configuration
            if not all([s3_config["s3_bucket"], s3_config["s3_access_key"], s3_config["s3_secret_key"], s3_config["s3_endpoint"]]):
                messagebox.showerror("Error", "Please fill in all S3 configuration fields")
                return
            
            # Save configuration
            self.db_manager.set_config("s3_config", json.dumps(s3_config))
            self.log("SUCCESS: S3 configuration saved")
            
            # Scan S3 for development images
            self.log("INFO: Scanning S3 for development images...")
            threading.Thread(target=self.scan_s3_for_dev_images, daemon=True).start()
            
        except Exception as e:
            self.log(f"ERROR: Failed to load S3 configuration: {e}")
            messagebox.showerror("Error", f"Failed to load S3 configuration: {e}")

    def scan_s3_for_dev_images(self):
        """Scan S3 repository for development images and populate UI from S3 metadata only"""
        try:
            # Load clients/sites directly from S3 metadata (not database)
            self.load_clients_from_s3_metadata()
            
            # Refresh the UI on main thread
            self.root.after(0, self.refresh_dev_ui_from_s3)
            
        except Exception as e:
            self.log(f"ERROR: Failed to scan S3 for development images: {e}")

    def load_clients_from_s3_metadata(self):
        """Load clients and sites from S3 metadata files in bucket root"""
        try:
            s3_config = {
                "s3_bucket": self.dev_s3_bucket_var.get(),
                "s3_access_key": self.dev_s3_access_var.get(),
                "s3_secret_key": self.dev_s3_secret_var.get(),
                "s3_endpoint": self.dev_s3_endpoint_var.get(),
                "s3_region": self.dev_s3_region_var.get()
            }
            
            if not all([s3_config["s3_bucket"], s3_config["s3_access_key"], s3_config["s3_secret_key"]]):
                self.log("WARNING: S3 configuration incomplete, cannot load metadata")
                return
            
            # Initialize storage for S3 metadata
            self.s3_clients = {}  # {client_uuid: {name, short_name, sites: {site_uuid: {name, short_name}}}}
            self.s3_images = {}   # {image_uuid: {client_uuid, site_uuid, role, status, created_date}}
            
            # Use boto3 to access S3 directly (fallback if AWS CLI not available)
            try:
                import boto3
                from botocore.exceptions import ClientError, NoCredentialsError
                
                # Create S3 client
                s3_client_kwargs = {
                    'aws_access_key_id': s3_config["s3_access_key"],
                    'aws_secret_access_key': s3_config["s3_secret_key"],
                    'region_name': s3_config["s3_region"]
                }
                
                # Add endpoint URL if not using AWS S3
                s3_endpoint = s3_config.get("s3_endpoint", "s3.amazonaws.com")
                if s3_endpoint != "s3.amazonaws.com":
                    s3_client_kwargs["endpoint_url"] = f"https://{s3_endpoint}"
                
                s3_client = boto3.client('s3', **s3_client_kwargs)
                
                # List all metadata files in bucket root /metadata/ folder
                try:
                    response = s3_client.list_objects_v2(
                        Bucket=s3_config["s3_bucket"],
                        Prefix="metadata/",
                        MaxKeys=1000
                    )
                    
                    if 'Contents' in response:
                        metadata_files = []
                        for obj in response['Contents']:
                            if obj['Key'].endswith('.json'):
                                metadata_files.append(obj['Key'])
                        
                        self.log(f"INFO: Found {len(metadata_files)} metadata files in S3")
                        
                        # Download and parse each metadata file
                        for metadata_file in metadata_files:
                            try:
                                # Download metadata file content
                                obj_response = s3_client.get_object(
                                    Bucket=s3_config["s3_bucket"],
                                    Key=metadata_file
                                )
                                
                                # Parse metadata
                                metadata_content = obj_response['Body'].read().decode('utf-8')
                                metadata = json.loads(metadata_content)
                                self.parse_s3_metadata(metadata)
                                
                            except Exception as e:
                                self.log(f"WARNING: Failed to process metadata file {metadata_file}: {e}")
                    else:
                        self.log(f"INFO: No metadata files found in S3 bucket /metadata/ folder")
                        
                except ClientError as e:
                    if e.response['Error']['Code'] == 'NoSuchBucket':
                        self.log(f"INFO: S3 bucket '{s3_config['s3_bucket']}' does not exist")
                    else:
                        self.log(f"ERROR: S3 access error: {e}")
                        
            except ImportError:
                self.log("ERROR: boto3 library not available. Please install: pip install boto3")
                return
            except NoCredentialsError:
                self.log("ERROR: Invalid S3 credentials")
                return
                
        except Exception as e:
            self.log(f"ERROR: Failed to load clients from S3 metadata: {e}")

    def parse_s3_metadata(self, metadata):
        """Parse individual S3 metadata file and extract client/site/image info"""
        try:
            tags = metadata.get('tags', {})
            
            # Extract information from tags
            client_uuid = tags.get('client-uuid')
            client_name = tags.get('client-name', 'Unknown Client')
            site_uuid = tags.get('site-uuid')
            site_name = tags.get('site-name', 'Unknown Site')
            image_uuid = metadata.get('backup_uuid')
            role = tags.get('role', 'Unknown')
            created_date = metadata.get('created_timestamp', '')
            
            # Determine image status based on metadata completeness
            status = "completed" if metadata.get('restic_snapshot_id') else "blank"
            
            if client_uuid and image_uuid:
                # Add client if not exists
                if client_uuid not in self.s3_clients:
                    # Create short name from client name if not provided
                    client_short = tags.get('client-short', client_name.upper().replace(' ', '')[:10])
                    self.s3_clients[client_uuid] = {
                        'name': client_name,
                        'short_name': client_short,
                        'sites': {}
                    }
                
                # Add site if not exists and site_uuid provided
                if site_uuid and site_uuid not in self.s3_clients[client_uuid]['sites']:
                    site_short = tags.get('site-short', site_name.upper().replace(' ', '')[:10])
                    self.s3_clients[client_uuid]['sites'][site_uuid] = {
                        'name': site_name,
                        'short_name': site_short
                    }
                
                # Add image
                self.s3_images[image_uuid] = {
                    'client_uuid': client_uuid,
                    'site_uuid': site_uuid,
                    'role': role,
                    'status': status,
                    'created_date': created_date
                }
                
        except Exception as e:
            self.log(f"WARNING: Failed to parse metadata: {e}")

    def refresh_dev_ui_from_s3(self):
        """Refresh development mode UI with data from S3 metadata"""
        try:
            # Populate client dropdown with S3 clients
            if hasattr(self, 's3_clients'):
                client_names = []
                for client_uuid, client_data in self.s3_clients.items():
                    display_name = f"{client_data['short_name']} ({client_data['name']})"
                    client_names.append(display_name)
                
                self.dev_client_combo['values'] = client_names
                self.log(f"INFO: Loaded {len(client_names)} clients from S3 metadata")
            else:
                self.dev_client_combo['values'] = []
                self.log("INFO: No clients found in S3 metadata")
                
        except Exception as e:
            self.log(f"ERROR: Failed to refresh development UI from S3: {e}")

    def refresh_dev_ui_from_db(self):
        """Refresh development mode UI with data from database"""
        try:
            # Get all clients from database
            clients = self.db_manager.get_clients()
            client_names = [f"{client[2]} ({client[1]})" for client in clients]  # short_name (name)
            
            self.dev_client_combo['values'] = client_names
            self.log(f"INFO: Loaded {len(clients)} clients for development mode")
            
        except Exception as e:
            self.log(f"ERROR: Failed to refresh development UI: {e}")

    def on_dev_client_selected(self, event=None):
        """Handle client selection in development mode - loads from S3 metadata"""
        try:
            selected = self.dev_client_var.get()
            if not selected:
                return
                
            # Extract client short name from selection
            client_short = selected.split(' (')[0]
            
            # Find client in S3 metadata
            client_uuid = None
            client_data = None
            
            if hasattr(self, 's3_clients'):
                for uuid, data in self.s3_clients.items():
                    if data['short_name'] == client_short:
                        client_uuid = uuid
                        client_data = data
                        break
            
            if client_data:
                # Load sites for this client from S3 metadata
                site_names = []
                for site_uuid, site_data in client_data['sites'].items():
                    display_name = f"{site_data['short_name']} ({site_data['name']})"
                    site_names.append(display_name)
                
                self.dev_site_combo['values'] = site_names
                self.log(f"INFO: Loaded {len(site_names)} sites for client {client_short} from S3")
                
                # Load development images for this client from S3 metadata
                self.load_dev_images_for_client_from_s3(client_uuid)
            else:
                self.log(f"WARNING: Client {client_short} not found in S3 metadata")
                
        except Exception as e:
            self.log(f"ERROR: Failed to load client data: {e}")

    def load_dev_images_for_client_from_s3(self, client_uuid):
        """Load development images for the selected client from S3 metadata"""
        try:
            # Clear current images list
            self.dev_images_listbox.delete(0, tk.END)
            
            if not hasattr(self, 's3_images') or not client_uuid:
                self.log("INFO: No images found for client")
                return
            
            # Find all images for this client
            client_images = []
            for image_uuid, image_data in self.s3_images.items():
                if image_data['client_uuid'] == client_uuid:
                    client_images.append((image_uuid, image_data))
            
            # Sort by created date (newest first)
            client_images.sort(key=lambda x: x[1]['created_date'], reverse=True)
            
            # Populate listbox with image info including status
            for image_uuid, image_data in client_images:
                created_date = image_data['created_date'][:10] if image_data['created_date'] else "Unknown"
                status = image_data['status'].upper()
                role = image_data['role']
                
                # Format: "Role - Status - Date - UUID"
                display_text = f"{role} - {status} - {created_date} - {image_uuid[:8]}"
                self.dev_images_listbox.insert(tk.END, display_text)
            
            self.log(f"INFO: Loaded {len(client_images)} development images for client")
            
        except Exception as e:
            self.log(f"ERROR: Failed to load development images from S3: {e}")

    def load_dev_images_for_client(self, client_id):
        """Load development images for the selected client"""
        try:
            # Get development images for this client
            images = self.db_manager.get_images_by_client_and_environment(client_id, "development")
            
            # Clear and populate listbox
            self.dev_images_listbox.delete(0, tk.END)
            
            for image in images:
                # Format: "Role - Date - UUID"
                created_date = image[4][:10] if image[4] else "Unknown"
                display_text = f"{image[3]} - {created_date} - {image[0][:8]}"
                self.dev_images_listbox.insert(tk.END, display_text)
            
            self.log(f"INFO: Loaded {len(images)} development images for client")
            
        except Exception as e:
            self.log(f"ERROR: Failed to load development images: {e}")

    def on_dev_image_selected(self, event=None):
        """Handle selection of an existing development image"""
        try:
            selection = self.dev_images_listbox.curselection()
            if selection:
                selected_text = self.dev_images_listbox.get(selection[0])
                # Extract UUID from the display text
                image_uuid = selected_text.split(' - ')[-1]
                self.log(f"INFO: Selected development image: {image_uuid}")
                
        except Exception as e:
            self.log(f"ERROR: Failed to handle image selection: {e}")

    def create_new_dev_client(self):
        """Create a new client for development mode - creates S3 metadata immediately"""
        try:
            # Simple dialog for client creation
            client_name = simpledialog.askstring("New Client", "Enter client name:")
            if not client_name:
                return
                
            client_short = simpledialog.askstring("New Client", "Enter client short name:")
            if not client_short:
                return
            
            # Create client and site together, then create blank image metadata
            site_name = simpledialog.askstring("New Client", "Enter initial site name:")
            if not site_name:
                site_name = f"{client_name} Main Site"
                
            site_short = simpledialog.askstring("New Client", "Enter site short name:")
            if not site_short:
                site_short = f"{client_short}MAIN"
            
            # Generate UUIDs
            client_uuid = generate_uuidv7()
            site_uuid = generate_uuidv7()
            image_uuid = generate_uuidv7()
            
            # Create blank image metadata and store to S3 immediately
            if self.create_blank_image_metadata_s3(client_uuid, client_name, client_short, 
                                                  site_uuid, site_name, site_short, image_uuid):
                self.log(f"SUCCESS: Created new client: {client_name} ({client_short}) with blank image")
                
                # Refresh from S3
                threading.Thread(target=self.scan_s3_for_dev_images, daemon=True).start()
                
                # Select the new client after a short delay for S3 refresh
                self.root.after(2000, lambda: self.select_created_client(client_short, client_name))
            else:
                messagebox.showerror("Error", "Failed to create client metadata in S3")
                
        except Exception as e:
            self.log(f"ERROR: Failed to create new client: {e}")
            messagebox.showerror("Error", f"Failed to create new client: {e}")

    def create_blank_image_metadata_s3(self, client_uuid, client_name, client_short, 
                                       site_uuid, site_name, site_short, image_uuid):
        """Create a blank image metadata file in S3 bucket root /metadata/ folder"""
        try:
            s3_config = {
                "s3_bucket": self.dev_s3_bucket_var.get(),
                "s3_access_key": self.dev_s3_access_var.get(),
                "s3_secret_key": self.dev_s3_secret_var.get(),
                "s3_endpoint": self.dev_s3_endpoint_var.get(),
                "s3_region": self.dev_s3_region_var.get()
            }
            
            # Create blank image metadata
            metadata = {
                "backup_uuid": image_uuid,
                "created_timestamp": datetime.now().isoformat(),
                "version": "1.0",
                "tool": "windows-image-prep-gui",
                "tool_version": "2025.1",
                "environment": "development",
                "status": "blank",
                "tags": {
                    "client-uuid": client_uuid,
                    "client-name": client_name,
                    "client-short": client_short,
                    "site-uuid": site_uuid,
                    "site-name": site_name,
                    "site-short": site_short,
                    "environment": "development",
                    "backup-uuid": image_uuid,
                    "created-date": datetime.now().isoformat(),
                    "role": "ADMIN"  # Default role for new clients
                }
            }
            
            # Upload metadata directly to S3 using boto3
            try:
                import boto3
                from botocore.exceptions import ClientError, NoCredentialsError
                
                # Create S3 client
                s3_client_kwargs = {
                    'aws_access_key_id': s3_config["s3_access_key"],
                    'aws_secret_access_key': s3_config["s3_secret_key"],
                    'region_name': s3_config["s3_region"]
                }
                
                # Add endpoint URL if not using AWS S3
                s3_endpoint = s3_config.get("s3_endpoint", "s3.amazonaws.com")
                if s3_endpoint != "s3.amazonaws.com":
                    s3_client_kwargs["endpoint_url"] = f"https://{s3_endpoint}"
                
                s3_client = boto3.client('s3', **s3_client_kwargs)
                
                # Convert metadata to JSON string
                metadata_json = json.dumps(metadata, indent=2)
                
                # Upload to S3 bucket root /metadata/ folder
                s3_key = f"metadata/{image_uuid}.json"
                
                s3_client.put_object(
                    Bucket=s3_config["s3_bucket"],
                    Key=s3_key,
                    Body=metadata_json.encode('utf-8'),
                    ContentType='application/json'
                )
                
                self.log(f"SUCCESS: Created blank image metadata in S3: {image_uuid}")
                return True
                
            except ImportError:
                self.log("ERROR: boto3 library not available. Please install: pip install boto3")
                return False
            except NoCredentialsError:
                self.log("ERROR: Invalid S3 credentials")
                return False
            except ClientError as e:
                self.log(f"ERROR: Failed to upload metadata to S3: {e}")
                return False
                
        except Exception as e:
            self.log(f"ERROR: Failed to create blank image metadata: {e}")
            return False

    def select_created_client(self, client_short, client_name):
        """Select the newly created client in the dropdown"""
        try:
            display_name = f"{client_short} ({client_name})"
            self.dev_client_var.set(display_name)
            self.on_dev_client_selected()
        except Exception as e:
            self.log(f"WARNING: Could not auto-select created client: {e}")

    def create_new_dev_site(self):
        """Create a new site for the selected client in development mode - creates S3 metadata"""
        try:
            # Check if client is selected
            if not self.dev_client_var.get():
                messagebox.showwarning("Warning", "Please select a client first")
                return
            
            # Get client info from S3 metadata
            client_short = self.dev_client_var.get().split(' (')[0]
            client_uuid = None
            client_name = None
            
            if hasattr(self, 's3_clients'):
                for uuid, data in self.s3_clients.items():
                    if data['short_name'] == client_short:
                        client_uuid = uuid
                        client_name = data['name']
                        break
            
            if not client_uuid:
                messagebox.showerror("Error", "Selected client not found in S3 metadata")
                return
            
            # Simple dialog for site creation
            site_name = simpledialog.askstring("New Site", "Enter site name:")
            if not site_name:
                return
                
            site_short = simpledialog.askstring("New Site", "Enter site short name:")
            if not site_short:
                return
            
            # Generate UUIDs
            site_uuid = generate_uuidv7()
            image_uuid = generate_uuidv7()
            
            # Create blank image metadata for new site
            if self.create_blank_image_metadata_s3(client_uuid, client_name, client_short, 
                                                  site_uuid, site_name, site_short, image_uuid):
                self.log(f"SUCCESS: Created new site: {site_name} ({site_short}) with blank image")
                
                # Refresh from S3
                threading.Thread(target=self.scan_s3_for_dev_images, daemon=True).start()
                
                # Select the new site after a short delay for S3 refresh
                self.root.after(2000, lambda: self.select_created_site(site_short, site_name))
            else:
                messagebox.showerror("Error", "Failed to create site metadata in S3")
                
        except Exception as e:
            self.log(f"ERROR: Failed to create new site: {e}")
            messagebox.showerror("Error", f"Failed to create new site: {e}")

    def select_created_site(self, site_short, site_name):
        """Select the newly created site in the dropdown"""
        try:
            display_name = f"{site_short} ({site_name})"
            self.dev_site_var.set(display_name)
        except Exception as e:
            self.log(f"WARNING: Could not auto-select created site: {e}")

    def create_dev_image(self):
        """Create a new development image"""
        try:
            # Validate selections
            if not all([self.dev_client_var.get(), self.dev_site_var.get(), self.dev_role_var.get()]):
                messagebox.showwarning("Warning", "Please select client, site, and role")
                return
            
            # Set development mode workflow  
            self.db_manager.set_config("workflow_mode", "development")
            
            # Start the backup process with development tagging
            self.start_dev_backup()
            
        except Exception as e:
            self.log(f"ERROR: Failed to create development image: {e}")
            messagebox.showerror("Error", f"Failed to create development image: {e}")

    def update_dev_image(self):
        """Update an existing development image"""
        try:
            # Check if an image is selected
            selection = self.dev_images_listbox.curselection()
            if not selection:
                messagebox.showwarning("Warning", "Please select an existing image to update")
                return
            
            # Validate role selection for update
            if not self.dev_role_var.get():
                messagebox.showwarning("Warning", "Please select a role for the image update")
                return
            
            # Set development mode workflow
            self.db_manager.set_config("workflow_mode", "development")
            
            # Start the backup process (this will update the existing repository)
            self.start_dev_backup()
            
        except Exception as e:
            self.log(f"ERROR: Failed to update development image: {e}")
            messagebox.showerror("Error", f"Failed to update development image: {e}")

    def start_dev_backup(self):
        """Start the development backup process"""
        try:
            self.log("INFO: Starting development image backup...")
            
            # Use existing backup functionality but with development tagging
            # This will call the same backup methods but with environment=development
            threading.Thread(target=self.perform_dev_backup_worker, daemon=True).start()
            
        except Exception as e:
            self.log(f"ERROR: Failed to start development backup: {e}")

    def perform_dev_backup_worker(self):
        """Worker thread for development backup"""
        try:
            # Extract client/site info
            client_short = self.dev_client_var.get().split(' (')[0]
            site_short = self.dev_site_var.get().split(' (')[0] if self.dev_site_var.get() else ""
            
            # Look up client in S3 metadata instead of database
            client_uuid = None
            client_name = None
            client_data = None
            
            if hasattr(self, 's3_clients'):
                for uuid, data in self.s3_clients.items():
                    if data['short_name'] == client_short:
                        client_uuid = uuid
                        client_name = data['name']
                        client_data = data
                        break
            
            if not client_uuid:
                self.log("ERROR: Client not found in S3 metadata")
                return
            
            # Look up site in S3 metadata if specified
            site_uuid = None
            site_name = None
            if site_short and client_data:
                for uuid, site_data in client_data['sites'].items():
                    if site_data['short_name'] == site_short:
                        site_uuid = uuid
                        site_name = site_data['name']
                        break
            
            # Build backup tags for development
            backup_tags = [
                f"client-uuid:{client_uuid}",
                f"client-name:{client_name}",
                f"environment:development",
                f"role:{self.dev_role_var.get()}",
                f"backup-uuid:{generate_uuidv7()}",
                f"created-date:{datetime.now().isoformat()}"
            ]
            
            if site_uuid:
                backup_tags.extend([
                    f"site-uuid:{site_uuid}",
                    f"site-name:{site_name}"
                ])
            
            # Add hardware info
            hardware_info = self.get_hardware_info()
            if hardware_info:
                for key, value in hardware_info.items():
                    if value:
                        backup_tags.append(f"hw-{key}:{value}")
            
            # Store backup tags for use by perform_restic_backup
            self._current_backup_tags = backup_tags
            
            # Download restic executable first
            restic_exe = self.download_restic()
            if not restic_exe:
                self.log("ERROR: Failed to download or locate restic.exe")
                return
            
            self.log(f"INFO: Using restic executable: {restic_exe}")
            
            # Perform the actual backup using existing restic functionality
            success = self.perform_restic_backup(restic_exe)
            
            if success:
                self.log("SUCCESS: Development image backup completed!")
                # Refresh the images list
                self.root.after(0, lambda: self.load_dev_images_for_client_from_s3(client_uuid))
            else:
                self.log("ERROR: Development image backup failed")
                
        except Exception as e:
            self.log(f"ERROR: Development backup worker failed: {e}")

    def create_production_capture_ui(self):
        """Create the PRODUCTION CAPTURE mode UI"""
        # Set repository type for production mode (can be S3 or local)
        self.repo_type_var = tk.StringVar(value="s3")
        
        # Create main frame for production capture mode
        production_frame = ttk.Frame(self.root)
        production_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self.mode_frames["production_capture"] = production_frame
        
        # Production mode notice
        notice_frame = ttk.LabelFrame(production_frame, text="Production Mode", padding="10")
        notice_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Label(notice_frame, text="⚠️ Production image capture creates deployment-ready images tagged as 'production'",
                 font=("TkDefaultFont", 10, "bold"), foreground="red").pack()
        
        # Similar structure to development mode but tagged as production
        # (This would be a simplified version focusing on production deployment)
        
        ttk.Label(production_frame, text="Production capture UI - Implementation pending",
                 font=("TkDefaultFont", 14)).pack(expand=True)

    def create_generalize_ui(self):
        """Create the GENERALIZE mode UI"""
        # Set default repository type
        self.repo_type_var = tk.StringVar(value="local")
        
        # Create main frame for generalize mode
        generalize_frame = ttk.Frame(self.root)
        generalize_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self.mode_frames["generalize"] = generalize_frame
        
        # Generalization notice
        notice_frame = ttk.LabelFrame(generalize_frame, text="System Generalization", padding="10")
        notice_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Label(notice_frame, text="🛠️ Prepare Windows images for deployment by running sysprep and cleanup tools",
                 font=("TkDefaultFont", 10, "bold")).pack()
        
        # Generalization tools would go here
        ttk.Label(generalize_frame, text="Generalization UI - Implementation pending",
                 font=("TkDefaultFont", 14)).pack(expand=True)

    def create_manage_images_ui(self):
        """Create the MANAGE IMAGES mode UI"""
        # Set default repository type
        self.repo_type_var = tk.StringVar(value="local")
        
        # Create main frame for manage images mode
        manage_frame = ttk.Frame(self.root)
        manage_frame.pack(fill="both", expand=True, padx=10, pady=5)
        self.mode_frames["manage_images"] = manage_frame
        
        # Image management notice
        notice_frame = ttk.LabelFrame(manage_frame, text="Image Management", padding="10")
        notice_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Label(notice_frame, text="📁 Browse, import, and manage existing images from local and S3 storage",
                 font=("TkDefaultFont", 10, "bold")).pack()
        
        # Image management tools would go here
        ttk.Label(manage_frame, text="Image management UI - Implementation pending",
                 font=("TkDefaultFont", 14)).pack(expand=True)

    def show_step(self, step_number):
        """Shows the specified step and hides others."""
        # Hide all step frames
        for frame in self.step_frames.values():
            frame.pack_forget()
        
        # Show current step frame
        if step_number in self.step_frames:
            self.step_frames[step_number].pack(fill="both", expand=True)
            self.current_step = step_number
            
            # Update header
            step_names = [
                "Create System Backup",
                "Professional Image & VM Management", 
                "Generalize & Cleanup",
                "Capture to WIM",
                "Deploy WIM"
            ]
            if self.current_step_label:
                self.current_step_label.config(text=f"Current Step: {step_number} - {step_names[step_number-1]}")
            
            # Update step button styles
            for i, button in self.step_buttons.items():
                if i == step_number:
                    # Current step - blue and bold
                    self.style.configure(f'Current.TButton', foreground='white', background='blue')
                    button.config(style='Current.TButton')
                elif i < step_number:
                    # Completed steps - green
                    self.style.configure(f'Completed.TButton', foreground='white', background='green')
                    button.config(style='Completed.TButton')
                else:
                    # Future steps - default style
                    button.config(style='TButton')
            
            # Update navigation buttons
            if self.prev_button:
                self.prev_button.config(state="normal" if step_number > 1 else "disabled")
            if self.next_button:
                self.next_button.config(state="normal" if step_number < self.total_steps else "disabled")
            
            # Log the navigation
            self.log(f"INFO: Navigated to Step {step_number} - {step_names[step_number-1]}")

    def previous_step(self):
        """Navigate to previous step."""
        if self.current_step > 1:
            self.show_step(self.current_step - 1)

    def next_step(self):
        """Navigate to next step."""
        if self.current_step < self.total_steps:
            self.show_step(self.current_step + 1)

    def populate_step1_frame(self):
        """Step 1: Create WIM Image"""
        frame = self.step_frames[1]
        
        # Step description
        desc_frame = ttk.LabelFrame(frame, text="Step 1: Create System Backup", padding="10")
        desc_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(desc_frame, text="Backup the current system using modern VSS + Restic method for maximum reliability.", 
                 font=("TkDefaultFont", 9)).pack()

        # Configuration
        config_frame = ttk.LabelFrame(frame, text="Configuration", padding="10")
        config_frame.pack(fill="x", pady=(0, 10))
        config_frame.columnconfigure(1, weight=1)

        # Repository Type Selection
        ttk.Label(config_frame, text="Repository Type:").grid(row=0, column=0, sticky="w", pady=2)
        self.repo_type_var = tk.StringVar(value="local")
        repo_type_frame = ttk.Frame(config_frame)
        repo_type_frame.grid(row=0, column=1, columnspan=2, sticky="w", pady=2)
        
        ttk.Radiobutton(repo_type_frame, text="Local File System", 
                       variable=self.repo_type_var, value="local",
                       command=self.on_repo_type_changed).pack(side="left", padx=(0, 20))
        ttk.Radiobutton(repo_type_frame, text="S3 Cloud Storage", 
                       variable=self.repo_type_var, value="s3",
                       command=self.on_repo_type_changed).pack(side="left")
        
        # Client and Site Selection (for development mode organization)
        self.client_site_frame = ttk.LabelFrame(config_frame, text="Client & Site Organization", padding="5")
        self.client_site_frame.grid(row=1, column=0, columnspan=3, sticky="we", pady=5)
        self.client_site_frame.columnconfigure(1, weight=1)
        self.client_site_frame.columnconfigure(3, weight=1)
        
        # Client selection
        ttk.Label(self.client_site_frame, text="Client:").grid(row=0, column=0, sticky="w", pady=2)
        self.client_var = tk.StringVar()
        self.client_combo = ttk.Combobox(self.client_site_frame, textvariable=self.client_var, state="readonly")
        self.client_combo.grid(row=0, column=1, sticky="we", padx=5)
        self.client_combo.bind('<<ComboboxSelected>>', self.on_client_selected)
        
        ttk.Button(self.client_site_frame, text="New Client", 
                  command=self.create_new_client, width=10).grid(row=0, column=2, padx=5)
        
        # Site selection
        ttk.Label(self.client_site_frame, text="Site:").grid(row=0, column=3, sticky="w", pady=2, padx=(20, 0))
        self.site_var = tk.StringVar()
        self.site_combo = ttk.Combobox(self.client_site_frame, textvariable=self.site_var, state="readonly")
        self.site_combo.grid(row=0, column=4, sticky="we", padx=5)
        
        ttk.Button(self.client_site_frame, text="New Site", 
                  command=self.create_new_site, width=10).grid(row=0, column=5, padx=5)
        
        # Role selection
        ttk.Label(self.client_site_frame, text="Role:").grid(row=1, column=0, sticky="w", pady=2)
        self.role_var = tk.StringVar(value="OP")
        role_combo = ttk.Combobox(self.client_site_frame, textvariable=self.role_var, 
                                 values=["ADMIN", "OP", "MANAGER", "VIP", "KIOSK", "SERVER", "IMAGING"])
        role_combo.grid(row=1, column=1, sticky="we", padx=5, pady=2)
        
        # Image selection (new vs existing)
        ttk.Label(self.client_site_frame, text="Image Type:").grid(row=2, column=0, sticky="w", pady=2)
        self.image_type_var = tk.StringVar(value="existing")
        image_type_frame = ttk.Frame(self.client_site_frame)
        image_type_frame.grid(row=2, column=1, columnspan=2, sticky="w", pady=2)
        
        ttk.Radiobutton(image_type_frame, text="New Image", 
                       variable=self.image_type_var, value="new",
                       command=self.on_image_type_changed).pack(side="left", padx=(0, 20))
        ttk.Radiobutton(image_type_frame, text="Update Existing", 
                       variable=self.image_type_var, value="existing",
                       command=self.on_image_type_changed).pack(side="left")
        
        # Existing image selection
        self.existing_image_frame = ttk.Frame(self.client_site_frame)
        self.existing_image_frame.grid(row=3, column=0, columnspan=6, sticky="we", pady=2)
        self.existing_image_frame.columnconfigure(1, weight=1)
        
        ttk.Label(self.existing_image_frame, text="Select Image:").grid(row=0, column=0, sticky="w", pady=2)
        self.existing_image_var = tk.StringVar()
        self.existing_image_combo = ttk.Combobox(self.existing_image_frame, textvariable=self.existing_image_var, 
                                               state="readonly", width=50)
        self.existing_image_combo.grid(row=0, column=1, sticky="we", padx=5)
        
        ttk.Button(self.existing_image_frame, text="Refresh", 
                  command=self.refresh_existing_images, width=10).grid(row=0, column=2, padx=5)
        
        # Initially hide existing image selection
        self.existing_image_frame.grid_remove()
        
        # Load initial client data
        self.refresh_client_site_data()
        
        # Scan S3 for existing images in development mode
        if self.get_workflow_mode() == "development":
            threading.Thread(target=self.scan_s3_for_images, daemon=True).start()
        
        # Only show client/site selection in development mode
        workflow_mode = self.get_workflow_mode()
        if workflow_mode != "development":
            self.client_site_frame.grid_remove()
        
        
        # S3 Configuration
        self.s3_config_frame = ttk.Frame(config_frame)
        self.s3_config_frame.grid(row=3, column=0, columnspan=3, sticky="we", pady=2)
        self.s3_config_frame.columnconfigure(1, weight=1)
        
        ttk.Label(self.s3_config_frame, text="S3 Bucket Path:").grid(row=0, column=0, sticky="w", pady=2)
        self.s3_path_var = tk.StringVar(value="")
        self.s3_path_entry = ttk.Entry(self.s3_config_frame, textvariable=self.s3_path_var)
        self.s3_path_entry.grid(row=0, column=1, sticky="we", padx=5)
        
        ttk.Button(self.s3_config_frame, text="Configure S3...", 
                  command=self.show_s3_configuration_dialog).grid(row=0, column=2, sticky="e")
        
        # S3 Status
        self.s3_status_var = tk.StringVar(value="S3 not configured")
        self.s3_status_label = ttk.Label(self.s3_config_frame, textvariable=self.s3_status_var, 
                                        font=("TkDefaultFont", 8), foreground="red")
        self.s3_status_label.grid(row=1, column=0, columnspan=3, sticky="w", pady=(2, 0))
        
        # Initialize repository type display and S3 status
        self.update_s3_status()
        self.on_repo_type_changed()
        
        # Network Credentials

        # Options
        options_frame = ttk.LabelFrame(frame, text="Capture Options", padding="10")
        options_frame.pack(fill="x", pady=(0, 10))
        
        self.capture_os_only_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Capture only Windows (OS) volume (Recommended)", 
                       variable=self.capture_os_only_var).pack(anchor="w")

        # Action Button
        button_frame = ttk.Frame(frame)
        button_frame.pack(pady=10, fill="x")
        
        # Create two buttons side by side
        button_container = ttk.Frame(button_frame)
        button_container.pack(fill="x")
        
        # VSS + Restic method (best)
        self.vss_create_button = ttk.Button(button_container, text="🚀 Create Backup with VSS + Restic (Recommended)", command=self.start_vss_restic_creation_thread)
        self.vss_create_button.pack(side="left", fill="x", expand=True, padx=(0, 5))
        
        # Direct DISM method (legacy/risky)
        self.direct_create_button = ttk.Button(button_container, text="⚠️ Legacy DISM Method (Risky)", command=self.start_direct_wim_creation_thread)
        self.direct_create_button.pack(side="right", fill="x", expand=True, padx=(5, 0))
        
        # Method explanation
        method_info = ttk.Label(frame, text="🚀 VSS + Restic: Modern backup tool, much more reliable than DISM\n⚠️ Legacy DISM: Old method with VSS timeout issues", 
                               font=("TkDefaultFont", 9), foreground="blue", justify="center")
        method_info.pack(pady=(5, 0))

    def populate_step2_frame(self):
        """Step 2: Professional Image & VM Management"""
        frame = self.step_frames[2]
        
        # Create notebook for tabbed interface
        self.step2_notebook = ttk.Notebook(frame)
        self.step2_notebook.pack(fill="both", expand=True, pady=5)
        
        # Tab 1: Create New Repository
        self.create_tab = ttk.Frame(self.step2_notebook)
        self.step2_notebook.add(self.create_tab, text="🆕 Create Repository")
        self.populate_create_image_tab()
        
        # Tab 2: Browse & Manage Repositories
        self.browse_tab = ttk.Frame(self.step2_notebook)
        self.step2_notebook.add(self.browse_tab, text="📁 Browse Repositories")
        self.populate_browse_images_tab()
        
        # Tab 3: Status Dashboard
        self.dashboard_tab = ttk.Frame(self.step2_notebook)
        self.step2_notebook.add(self.dashboard_tab, text="📊 Dashboard")
        self.populate_dashboard_tab()
        
        # Tab 4: Database Management
        self.database_tab = ttk.Frame(self.step2_notebook)
        self.step2_notebook.add(self.database_tab, text="🗄️ Database")
        self.populate_database_tab()
        
        # Add log area below the notebook
        log_frame = ttk.LabelFrame(frame, text="Progress & Log Output", padding="5")
        log_frame.pack(fill="both", expand=True, pady=(10, 0))
        
        # Create step 2 specific log text widget
        self.step2_log_text = scrolledtext.ScrolledText(
            log_frame, 
            height=8, 
            font=("Consolas", 9), 
            bg="#1e1e1e", 
            fg="#ffffff", 
            insertbackground="#ffffff"
        )
        self.step2_log_text.pack(fill="both", expand=True)
        
        # Add initial message
        self.step2_log_text.insert(tk.END, "[INFO] Step 2: Professional Image & VM Management loaded\n")
        self.step2_log_text.insert(tk.END, "[INFO] Repository base path: " + str(self.get_restic_base_path()) + "\n")
        self.step2_log_text.insert(tk.END, "[INFO] Working VHDX directory: " + str(self.db.get_working_vhdx_directory()) + "\n\n")
        self.step2_log_text.see(tk.END)

    def log_step2(self, message):
        """Log message to Step 2 log area"""
        if hasattr(self, 'step2_log_text'):
            timestamp = datetime.now().strftime("%H:%M:%S")
            formatted_message = f"[{timestamp}] {message}\n"
            
            self.step2_log_text.insert(tk.END, formatted_message)
            self.step2_log_text.see(tk.END)
            self.step2_log_text.update()
        
        # Also log to main log
        self.log(message)

    def populate_create_image_tab(self):
        """Populate the Create New Image tab"""
        frame = self.create_tab
        
        # Step description
        desc_frame = ttk.LabelFrame(frame, text="Create New Restic Repository", padding="10")
        desc_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(desc_frame, text="Create a new restic repository for client system backups and images", 
                 font=("TkDefaultFont", 9)).pack()

        # Client/Site Configuration
        client_frame = ttk.LabelFrame(frame, text="Client & Site Information", padding="10")
        client_frame.pack(fill="x", pady=(0, 10))
        client_frame.columnconfigure(1, weight=1)
        
        # Client selection/creation
        ttk.Label(client_frame, text="Client:").grid(row=0, column=0, sticky="w", pady=2)
        client_container = ttk.Frame(client_frame)
        client_container.grid(row=0, column=1, sticky="we", padx=5)
        client_container.columnconfigure(0, weight=1)
        
        self.client_var = tk.StringVar()
        self.client_combo = ttk.Combobox(client_container, textvariable=self.client_var, width=30)
        self.client_combo.grid(row=0, column=0, sticky="we", padx=(0, 5))
        self.client_combo.bind('<<ComboboxSelected>>', self.on_client_selected)
        
        ttk.Button(client_container, text="New Client", command=self.create_new_client, width=12).grid(row=0, column=1)
        
        # Site selection/creation
        ttk.Label(client_frame, text="Site:").grid(row=1, column=0, sticky="w", pady=2)
        site_container = ttk.Frame(client_frame)
        site_container.grid(row=1, column=1, sticky="we", padx=5)
        site_container.columnconfigure(0, weight=1)
        
        self.site_var = tk.StringVar()
        self.site_combo = ttk.Combobox(site_container, textvariable=self.site_var, width=30)
        self.site_combo.grid(row=0, column=0, sticky="we", padx=(0, 5))
        
        ttk.Button(site_container, text="New Site", command=self.create_new_site, width=12).grid(row=0, column=1)
        
        # Role
        ttk.Label(client_frame, text="Image Role:").grid(row=2, column=0, sticky="w", pady=2)
        self.role_var = tk.StringVar()
        role_combo = ttk.Combobox(client_frame, textvariable=self.role_var, values=[
            "Desktop", "Server", "Workstation", "Domain Controller", "Database", 
            "Web Server", "File Server", "Terminal Server", "Custom"
        ])
        role_combo.grid(row=2, column=1, sticky="we", padx=5)
        role_combo.set("Desktop")

        # Repository Configuration
        config_frame = ttk.LabelFrame(frame, text="Repository Configuration", padding="10")
        config_frame.pack(fill="x", pady=(0, 10))
        config_frame.columnconfigure(1, weight=1)

        # Repository Name (auto-generated from client)
        ttk.Label(config_frame, text="Repository Name:").grid(row=0, column=0, sticky="w", pady=2)
        self.repo_name_var = tk.StringVar()
        self.repo_name_entry = ttk.Entry(config_frame, textvariable=self.repo_name_var, state="readonly")
        self.repo_name_entry.grid(row=0, column=1, sticky="we", padx=5)
        
        # Repository Location
        ttk.Label(config_frame, text="Repository Location:").grid(row=1, column=0, sticky="w", pady=2)
        self.repo_location_var = tk.StringVar()
        self.repo_location_entry = ttk.Entry(config_frame, textvariable=self.repo_location_var)
        self.repo_location_entry.grid(row=1, column=1, sticky="we", padx=5)
        ttk.Button(config_frame, text="Browse...", command=self.browse_repo_location).grid(row=1, column=2)
        
        # Repository Password
        ttk.Label(config_frame, text="Repository Password:").grid(row=2, column=0, sticky="w", pady=2)
        self.repo_password_var = tk.StringVar(value="SystemBackup2024!")
        self.repo_password_entry = ttk.Entry(config_frame, textvariable=self.repo_password_var, show="*")
        self.repo_password_entry.grid(row=2, column=1, sticky="we", padx=5)
        
        # Auto-generate password button
        ttk.Button(config_frame, text="Generate", command=self.generate_repo_password).grid(row=2, column=2)
        
        # Import existing repository option
        import_frame = ttk.LabelFrame(frame, text="Import Existing Repository", padding="10")
        import_frame.pack(fill="x", pady=(0, 10))
        import_frame.columnconfigure(1, weight=1)
        
        self.import_existing_var = tk.BooleanVar()
        ttk.Checkbutton(import_frame, text="Import existing restic repository", 
                       variable=self.import_existing_var, 
                       command=self.toggle_import_mode).grid(row=0, column=0, columnspan=3, sticky="w")
        
        ttk.Label(import_frame, text="Existing Repository:").grid(row=1, column=0, sticky="w", pady=2)
        self.import_repo_var = tk.StringVar()
        self.import_repo_entry = ttk.Entry(import_frame, textvariable=self.import_repo_var, state="disabled")
        self.import_repo_entry.grid(row=1, column=1, sticky="we", padx=5)
        self.import_browse_btn = ttk.Button(import_frame, text="Browse...", 
                                          command=self.browse_import_repo, state="disabled")
        self.import_browse_btn.grid(row=1, column=2)
        
        # Import button
        self.import_repo_btn = ttk.Button(import_frame, text="📥 Import Repository", 
                                        command=self.import_selected_repository, state="disabled")
        self.import_repo_btn.grid(row=2, column=1, pady=10, sticky="ew")

        # VM Configuration
        vm_frame = ttk.LabelFrame(frame, text="Virtual Machine Configuration", padding="10")
        vm_frame.pack(fill="x", pady=(0, 10))
        vm_frame.columnconfigure(1, weight=1)
        
        self.create_vm_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(vm_frame, text="Create VM after image creation", 
                       variable=self.create_vm_var).grid(row=0, column=0, columnspan=3, sticky="w")
        
        ttk.Label(vm_frame, text="RAM (GB):").grid(row=1, column=0, sticky="w", pady=2)
        self.vm_ram_var = tk.IntVar(value=4)
        ttk.Spinbox(vm_frame, from_=1, to=64, textvariable=self.vm_ram_var, width=8).grid(row=1, column=1, sticky="w", padx=5)
        
        ttk.Label(vm_frame, text="CPUs:").grid(row=1, column=2, sticky="w", pady=2, padx=(20, 0))
        self.vm_cpu_var = tk.IntVar(value=4)
        ttk.Spinbox(vm_frame, from_=1, to=16, textvariable=self.vm_cpu_var, width=8).grid(row=1, column=3, sticky="w", padx=5)

        # Action Button
        self.create_image_button = ttk.Button(frame, text="🚀 Create Professional Image & VM", 
                                            command=self.start_professional_image_creation)
        self.create_image_button.pack(pady=20, fill="x")
        
        # Load initial data
        self.refresh_client_site_data()
        
        # Update dashboard when tab changes
        self.step2_notebook.bind("<<NotebookTabChanged>>", self.on_step2_tab_changed)

    def populate_browse_images_tab(self):
        """Populate the Browse Images tab"""
        frame = self.browse_tab
        
        # Controls frame
        controls_frame = ttk.Frame(frame)
        controls_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Button(controls_frame, text="🔄 Refresh", command=self.refresh_images_list).pack(side="left", padx=(0, 10))
        ttk.Button(controls_frame, text="🔍 Scan & Import Repository", command=self.scan_and_import_repository).pack(side="left", padx=(0, 10))
        ttk.Button(controls_frame, text="💿 Create VHDX", command=self.create_vhdx_dialog).pack(side="left", padx=(0, 10))
        ttk.Button(controls_frame, text="💾 Restore Selected", command=self.restore_selected_repository).pack(side="left", padx=(0, 10))
        ttk.Button(controls_frame, text="📥 Import External Repository", command=self.import_repository_dialog).pack(side="left", padx=(0, 10))
        ttk.Button(controls_frame, text="🔍 Check for Orphans", command=self.check_orphan_files).pack(side="left", padx=(0, 10))
        
        # Repositories list
        images_frame = ttk.LabelFrame(frame, text="Existing Repositories", padding="10")
        images_frame.pack(fill="both", expand=True)
        
        # Treeview for repositories
        columns = ("Client", "Site", "Role", "Size", "Snapshots", "Type", "Created", "Status")
        self.images_tree = ttk.Treeview(images_frame, columns=columns, show="headings", height=15)
        
        for col in columns:
            self.images_tree.heading(col, text=col)
            self.images_tree.column(col, width=100)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(images_frame, orient="vertical", command=self.images_tree.yview)
        h_scrollbar = ttk.Scrollbar(images_frame, orient="horizontal", command=self.images_tree.xview)
        self.images_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        self.images_tree.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        
        images_frame.columnconfigure(0, weight=1)
        images_frame.rowconfigure(0, weight=1)
        
        # Context menu for images
        self.images_tree.bind("<Button-3>", self.show_image_context_menu)
        self.images_tree.bind("<Double-1>", self.on_image_double_click)

    def populate_dashboard_tab(self):
        """Populate the Dashboard tab"""
        frame = self.dashboard_tab
        
        # Stats frame
        stats_frame = ttk.LabelFrame(frame, text="Statistics", padding="10")
        stats_frame.pack(fill="x", pady=(0, 10))
        
        stats_container = ttk.Frame(stats_frame)
        stats_container.pack(fill="x")
        
        self.stats_labels = {}
        stats = ["Total Images", "Total VMs", "Total Clients", "Total Sites", "Storage Used"]
        
        for i, stat in enumerate(stats):
            col = i % 3
            row = i // 3
            
            stat_frame = ttk.Frame(stats_container)
            stat_frame.grid(row=row, column=col, padx=10, pady=5, sticky="w")
            
            ttk.Label(stat_frame, text=f"{stat}:", font=("TkDefaultFont", 9, "bold")).pack(anchor="w")
            label = ttk.Label(stat_frame, text="Loading...", font=("TkDefaultFont", 12))
            label.pack(anchor="w")
            self.stats_labels[stat] = label
        
        # Recent activity
        activity_frame = ttk.LabelFrame(frame, text="Recent Activity", padding="10")
        activity_frame.pack(fill="both", expand=True)
        
        self.activity_tree = ttk.Treeview(activity_frame, columns=("Time", "Action", "Details"), show="headings")
        self.activity_tree.heading("Time", text="Time")
        self.activity_tree.heading("Action", text="Action") 
        self.activity_tree.heading("Details", text="Details")
        self.activity_tree.pack(fill="both", expand=True)

    def populate_database_tab(self):
        """Populate the Database Management tab"""
        frame = self.database_tab
        
        # Database info
        info_frame = ttk.LabelFrame(frame, text="Database Information", padding="10")
        info_frame.pack(fill="x", pady=(0, 10))
        
        db_path_label = ttk.Label(info_frame, text=f"Database Location: {self.db.db_path}")
        db_path_label.pack(anchor="w")
        
        storage_label = ttk.Label(info_frame, text=f"Image Storage: {self.image_store_path}")
        storage_label.pack(anchor="w")
        
        # Database operations
        ops_frame = ttk.LabelFrame(frame, text="Database Operations", padding="10")
        ops_frame.pack(fill="x", pady=(0, 10))
        
        button_row1 = ttk.Frame(ops_frame)
        button_row1.pack(fill="x", pady=5)
        
        ttk.Button(button_row1, text="📤 Export Database", command=self.export_database).pack(side="left", padx=(0, 10))
        ttk.Button(button_row1, text="📥 Import Database", command=self.import_database).pack(side="left", padx=(0, 10))
        ttk.Button(button_row1, text="🔄 Backup Database", command=self.backup_database).pack(side="left")
        
        button_row2 = ttk.Frame(ops_frame)
        button_row2.pack(fill="x", pady=5)
        
        ttk.Button(button_row2, text="🧹 Clean Orphaned Records", command=self.clean_orphaned_records).pack(side="left", padx=(0, 10))
        ttk.Button(button_row2, text="📊 Database Statistics", command=self.show_database_stats).pack(side="left")
        
        # Client/Site management
        mgmt_frame = ttk.LabelFrame(frame, text="Client & Site Management", padding="10")
        mgmt_frame.pack(fill="both", expand=True)
        
        # Split into two columns
        left_col = ttk.Frame(mgmt_frame)
        left_col.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        right_col = ttk.Frame(mgmt_frame)
        right_col.pack(side="right", fill="both", expand=True)
        
        # Clients list
        ttk.Label(left_col, text="Clients", font=("TkDefaultFont", 10, "bold")).pack(anchor="w")
        self.clients_tree = ttk.Treeview(left_col, columns=("Name", "Short Name"), show="headings", height=10)
        self.clients_tree.heading("Name", text="Name")
        self.clients_tree.heading("Short Name", text="Short Name")
        self.clients_tree.pack(fill="both", expand=True, pady=(5, 0))
        
        # Sites list
        ttk.Label(right_col, text="Sites", font=("TkDefaultFont", 10, "bold")).pack(anchor="w")
        self.sites_tree = ttk.Treeview(right_col, columns=("Client", "Name", "Short Name"), show="headings", height=10)
        self.sites_tree.heading("Client", text="Client")
        self.sites_tree.heading("Name", text="Name")
        self.sites_tree.heading("Short Name", text="Short Name")
        self.sites_tree.pack(fill="both", expand=True, pady=(5, 0))

    # === Step 2 Methods ===
    
    def on_client_selected(self, event=None):
        """Handle client selection to update sites list"""
        try:
            client_name = self.client_var.get()
            print(f"DEBUG: on_client_selected called with client_name: '{client_name}'")
            
            if not client_name:
                print("DEBUG: No client name, returning")
                return
            
            # Find client ID
            clients = self.db.get_clients()
            print(f"DEBUG: Found {len(clients)} clients in database")
            
            client_id = None
            for cid, name, short_name, desc in clients:
                print(f"DEBUG: Checking client: '{name}' (ID: {cid})")
                if name == client_name:
                    client_id = cid
                    print(f"DEBUG: Found matching client ID: {client_id}")
                    break
            
            if client_id:
                sites = self.db.get_sites(client_id)
                print(f"DEBUG: Found {len(sites)} sites for client {client_id}")
                
                site_names = [name for _, _, name, _, _, _ in sites]
                print(f"DEBUG: Site names: {site_names}")
                
                # Store the sites for reference and handle commas properly
                self.current_client_sites = sites
                
                # Check if site_combo exists
                if hasattr(self, 'site_combo'):
                    # Use tuple format to handle commas in site names
                    self.site_combo['values'] = tuple(site_names)
                    if site_names:
                        self.site_var.set(site_names[0])  # Set the StringVar directly
                        print(f"DEBUG: Set site variable to: {site_names[0]}")
                    else:
                        self.site_var.set("")  # Clear if no sites
                        print("DEBUG: No sites available")
                else:
                    print("DEBUG: site_combo attribute not found")
            else:
                print(f"DEBUG: No client ID found for '{client_name}'")
                    
            # Update repository name based on client selection
            if hasattr(self, 'update_repo_name'):
                self.update_repo_name()
            else:
                print("DEBUG: update_repo_name method not found")
                
        except Exception as e:
            print(f"DEBUG: Exception in on_client_selected: {e}")
            if hasattr(self, 'log'):
                self.log(f"ERROR: Failed to update sites list: {e}")
            else:
                print(f"ERROR: Failed to update sites list: {e}")

    def create_new_client(self):
        """Create new client dialog"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Create New Client")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # Form fields
        ttk.Label(dialog, text="Client Name:").pack(pady=5)
        name_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=name_var, width=40).pack(pady=5)
        
        ttk.Label(dialog, text="Short Name (for VM naming):").pack(pady=5)
        short_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=short_var, width=40).pack(pady=5)
        
        ttk.Label(dialog, text="Description (optional):").pack(pady=5)
        desc_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=desc_var, width=40).pack(pady=5)
        
        def save_client():
            name = name_var.get().strip()
            short = short_var.get().strip()
            desc = desc_var.get().strip()
            
            if not name or not short:
                messagebox.showerror("Error", "Name and Short Name are required")
                return
            
            try:
                self.db.add_client(name, short, desc)
                self.refresh_client_site_data()
                self.client_var.set(name)
                self.log(f"INFO: Created new client: {name}")
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to create client: {e}")
        
        ttk.Button(dialog, text="Create Client", command=save_client).pack(pady=20)

    def create_new_site(self):
        """Create new site dialog"""
        client_name = self.client_var.get().strip()
        print(f"DEBUG: create_new_site - client_name: '{client_name}'")
        
        if not client_name or client_name == "-- Select Client --":
            messagebox.showerror("Error", "Please select a client first")
            return
        
        # Find client ID
        clients = self.db.get_clients()
        client_id = None
        for cid, name, short_name, desc in clients:
            if name == client_name:
                client_id = cid
                break
        
        if not client_id:
            messagebox.showerror("Error", "Selected client not found")
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Create New Site for {client_name}")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # Form fields
        ttk.Label(dialog, text="Site Name:").pack(pady=5)
        name_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=name_var, width=40).pack(pady=5)
        
        ttk.Label(dialog, text="Short Name (for VM naming):").pack(pady=5)
        short_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=short_var, width=40).pack(pady=5)
        
        ttk.Label(dialog, text="Description (optional):").pack(pady=5)
        desc_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=desc_var, width=40).pack(pady=5)
        
        def save_site():
            name = name_var.get().strip()
            short = short_var.get().strip()
            desc = desc_var.get().strip()
            
            if not name or not short:
                messagebox.showerror("Error", "Name and Short Name are required")
                return
            
            try:
                # Find client ID
                clients = self.db.get_clients()
                client_id = None
                for cid, n, _, _ in clients:
                    if n == self.client_var.get():
                        client_id = cid
                        break
                
                if client_id:
                    self.db.add_site(client_id, name, short, desc)
                    # Refresh the parent dialog's site combo
                    self.refresh_client_site_data()
                    # Auto-select the newly created site
                    self.site_var.set(name)
                    # Trigger site selection event
                    self.on_client_selected()
                    self.log(f"SUCCESS: Created new site: {name}")
                    dialog.destroy()
                else:
                    messagebox.showerror("Error", "Could not create site: Client ID not found.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to create site: {e}")
        
        ttk.Button(dialog, text="Create Site", command=save_site).pack(pady=20)

    def browse_wim_source(self):
        """Browse for WIM source file"""
        path = filedialog.askopenfilename(
            title="Select Source WIM File",
            filetypes=[("WIM Files", "*.wim"), ("All Files", "*.*")]
        )
        if path:
            self.wim_source_var.set(path)

    def browse_repo_location(self):
        """Browse for repository location"""
        path = filedialog.askdirectory(
            title="Select Repository Location"
        )
        if path:
            self.repo_location_var.set(path)
            
    def browse_import_repo(self):
        """Browse for existing repository to import"""
        path = filedialog.askdirectory(
            title="Select Existing Restic Repository"
        )
        if path:
            self.import_repo_var.set(path)
            
    def generate_repo_password(self):
        """Generate a random repository password"""
        import secrets
        import string
        alphabet = string.ascii_letters + string.digits + "!@#$%^&*"
        password = ''.join(secrets.choice(alphabet) for _ in range(16))
        self.repo_password_var.set(password)
        
    def toggle_import_mode(self):
        """Toggle between create new and import existing repository"""
        if self.import_existing_var.get():
            # Enable import controls
            self.import_repo_entry.config(state="normal")
            self.import_browse_btn.config(state="normal")
            self.import_repo_btn.config(state="normal")
            # Disable create controls
            self.repo_location_entry.config(state="disabled")
            self.repo_password_entry.config(state="disabled")
        else:
            # Enable create controls
            self.repo_location_entry.config(state="normal")
            self.repo_password_entry.config(state="normal")
            # Disable import controls
            self.import_repo_entry.config(state="disabled")
            self.import_browse_btn.config(state="disabled")
            self.import_repo_btn.config(state="disabled")
            
    def import_selected_repository(self):
        """Import the selected repository to organized client directory structure"""
        try:
            # Get form data
            source_repo = self.import_repo_var.get().strip()
            client_name = self.client_var.get().strip()
            site_name = self.site_var.get().strip()
            role = self.role_var.get().strip()
            
            # Validation
            if not source_repo or not Path(source_repo).exists():
                messagebox.showerror("Error", "Please select a valid repository to import")
                return
            
            if not client_name:
                messagebox.showerror("Error", "Please select a client")
                return
            
            if not site_name:
                messagebox.showerror("Error", "Please select a site")
                return
            
            # Get client and site IDs
            clients = self.db.get_clients()
            client_id = None
            for cid, name, _, _ in clients:
                if name == client_name:
                    client_id = cid
                    break
            
            if not client_id:
                messagebox.showerror("Error", f"Client '{client_name}' not found")
                return
            
            sites = self.db.get_sites(client_id)
            site_id = None
            for sid, _, name, _, _, _ in sites:
                if name == site_name:
                    site_id = sid
                    break
            
            if not site_id:
                messagebox.showerror("Error", f"Site '{site_name}' not found")
                return
            
            # Prompt for repository password
            repo_password = simpledialog.askstring("Repository Password", 
                                                  "Enter the password for the repository to import:",
                                                  show='*')
            if not repo_password:
                messagebox.showerror("Error", "Repository password is required")
                return
            
            # Confirm import
            if not messagebox.askyesno("Confirm Import", 
                                     f"Import repository to organized structure?\n\n"
                                     f"Source: {source_repo}\n"
                                     f"Client: {client_name}\n"
                                     f"Site: {site_name}\n"
                                     f"Role: {role}\n\n"
                                     "The repository will be copied to:\n"
                                     f"{self.get_restic_base_path()}\\{client_id}\n\n"
                                     "Continue?"):
                return
            
            # Start import with progress dialog
            self.import_repo_btn.config(state="disabled")
            self.log_step2(f"Starting import of repository: {source_repo}")
            
            # Create progress dialog
            self.show_import_progress_dialog(source_repo, client_id, site_id, role, repo_password)
            
        except Exception as e:
            self.log_step2(f"Import setup failed: {str(e)}")
            messagebox.showerror("Error", f"Failed to start import: {e}")

    def show_import_progress_dialog(self, source_repo, client_id, site_id, role, password):
        """Show progress dialog during repository import"""
        progress_dialog = tk.Toplevel(self.root)
        progress_dialog.title("Importing Repository")
        progress_dialog.geometry("800x600")
        progress_dialog.transient(self.root)
        progress_dialog.grab_set()
        
        # Center dialog
        progress_dialog.update_idletasks()
        x = (progress_dialog.winfo_screenwidth() // 2) - (progress_dialog.winfo_width() // 2)
        y = (progress_dialog.winfo_screenheight() // 2) - (progress_dialog.winfo_height() // 2)
        progress_dialog.geometry(f"+{x}+{y}")
        
        # Title
        ttk.Label(progress_dialog, text="Repository Import Progress", 
                 font=("TkDefaultFont", 14, "bold")).pack(pady=10)
        
        # Status
        status_var = tk.StringVar(value="Preparing import...")
        status_label = ttk.Label(progress_dialog, textvariable=status_var, font=("TkDefaultFont", 10))
        status_label.pack(pady=5)
        
        # Progress bar
        progress_frame = ttk.Frame(progress_dialog)
        progress_frame.pack(fill="x", padx=20, pady=10)
        
        progress_bar = ttk.Progressbar(progress_frame, mode='indeterminate')
        progress_bar.pack(fill="x")
        progress_bar.start()
        
        # Log output
        log_frame = ttk.LabelFrame(progress_dialog, text="Import Log", padding="10")
        log_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        log_text = scrolledtext.ScrolledText(
            log_frame, 
            height=20, 
            font=("Consolas", 9),
            bg="#1e1e1e", 
            fg="#ffffff", 
            insertbackground="#ffffff"
        )
        log_text.pack(fill="both", expand=True)
        
        # Close button (initially disabled)
        button_frame = ttk.Frame(progress_dialog)
        button_frame.pack(fill="x", padx=20, pady=10)
        
        close_btn = ttk.Button(button_frame, text="Close", command=progress_dialog.destroy, state="disabled")
        close_btn.pack(side="right")
        
        cancel_btn = ttk.Button(button_frame, text="Cancel", state="disabled")  # We'll implement cancel later
        cancel_btn.pack(side="right", padx=(0, 10))
        
        def log_to_dialog(message):
            """Log message to the dialog"""
            timestamp = datetime.now().strftime("%H:%M:%S")
            formatted_message = f"[{timestamp}] {message}\n"
            log_text.insert(tk.END, formatted_message)
            log_text.see(tk.END)
            log_text.update()
            # Also log to Step 2
            self.log_step2(message)
        
        # Start import in thread
        def import_thread():
            try:
                log_to_dialog("Starting repository import...")
                status_var.set("Importing repository...")
                
                success = self.perform_repository_import_with_logging(
                    source_repo, client_id, site_id, role, password, log_to_dialog, status_var
                )
                
                progress_dialog.after(0, lambda: import_complete(success))
                
            except Exception as e:
                progress_dialog.after(0, lambda: import_failed(str(e)))
        
        def import_complete(success):
            progress_bar.stop()
            close_btn.config(state="normal")
            cancel_btn.config(state="disabled")
            self.import_repo_btn.config(state="normal")
            
            if success:
                status_var.set("Import completed successfully!")
                log_to_dialog("✓ Repository import completed successfully!")
                # Clear the import path
                self.import_repo_var.set("")
                # Refresh Step 2 data
                self.refresh_images_list()
            else:
                status_var.set("Import failed!")
                log_to_dialog("✗ Repository import failed!")
        
        def import_failed(error):
            progress_bar.stop()
            close_btn.config(state="normal")
            cancel_btn.config(state="disabled")
            self.import_repo_btn.config(state="normal")
            status_var.set("Import failed!")
            log_to_dialog(f"✗ FATAL ERROR: {error}")
        
        threading.Thread(target=import_thread, daemon=True).start()

    def perform_repository_import_with_logging(self, source_repo, client_id, site_id, role, password, log_func, status_var):
        """Perform repository import with detailed logging"""
        try:
            source_path = Path(source_repo)
            log_func(f"Source repository: {source_path}")
            
            # Create destination path: {restic_base}/{client_uuid} (no subfolder)
            restic_base = self.get_restic_base_path()
            dest_repo = restic_base / client_id
            
            log_func(f"Destination path: {dest_repo}")
            log_func(f"Client ID: {client_id}")
            
            # Check if destination already exists and has repository files
            if dest_repo.exists():
                existing_items = list(dest_repo.iterdir())
                if existing_items:
                    log_func(f"ERROR: Client directory already contains {len(existing_items)} items:")
                    for item in existing_items[:5]:  # Show first 5 items
                        log_func(f"  - {item.name}")
                    if len(existing_items) > 5:
                        log_func(f"  ... and {len(existing_items) - 5} more items")
                    raise Exception(f"Client directory already contains files. Cannot import repository.")
            
            # Create client directory if needed
            log_func("Creating destination directory...")
            dest_repo.mkdir(parents=True, exist_ok=True)
            log_func(f"✓ Created client directory: {dest_repo}")
            
            # Get source items
            source_items = list(source_path.iterdir())
            log_func(f"Found {len(source_items)} items to copy from source")
            
            # Copy repository contents (not the folder itself)
            log_func("Copying repository files...")
            status_var.set("Copying repository files...")
            
            copied_count = 0
            for item in source_items:
                log_func(f"Copying: {item.name}")
                try:
                    if item.is_dir():
                        shutil.copytree(item, dest_repo / item.name)
                        log_func(f"✓ Copied directory: {item.name}")
                    else:
                        shutil.copy2(item, dest_repo / item.name)
                        log_func(f"✓ Copied file: {item.name}")
                    copied_count += 1
                except Exception as e:
                    log_func(f"✗ Failed to copy {item.name}: {e}")
                    raise Exception(f"Failed to copy {item.name}: {e}")
            
            log_func(f"✓ Successfully copied {copied_count} items")
            
            # Verify repository integrity
            log_func("Verifying repository integrity...")
            status_var.set("Verifying repository...")
            
            restic_exe = self.download_restic()
            if not restic_exe:
                log_func("✗ Could not obtain restic binary")
                raise Exception("Could not obtain restic binary")
            
            log_func(f"✓ Using restic: {restic_exe}")
            
            # Test repository access
            log_func("Setting up restic environment...")
            os.environ['RESTIC_REPOSITORY'] = str(dest_repo)
            os.environ['RESTIC_PASSWORD'] = password
            log_func(f"RESTIC_REPOSITORY = {dest_repo}")
            
            log_func("Testing repository access...")
            check_cmd = [restic_exe, "snapshots", "--json"]
            log_func(f"Running command: {' '.join(check_cmd)}")
            
            result = subprocess.run(check_cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
            
            log_func(f"Command return code: {result.returncode}")
            if result.stdout:
                log_func(f"STDOUT: {result.stdout[:200]}..." if len(result.stdout) > 200 else f"STDOUT: {result.stdout}")
            if result.stderr:
                log_func(f"STDERR: {result.stderr}")
            
            if result.returncode != 0:
                log_func("✗ Repository verification failed!")
                log_func("Cleaning up copied files...")
                # Clean up on failure - remove all copied files from client directory
                for item in dest_repo.iterdir():
                    try:
                        if item.is_dir():
                            shutil.rmtree(item)
                            log_func(f"✓ Removed directory: {item.name}")
                        else:
                            item.unlink()
                            log_func(f"✓ Removed file: {item.name}")
                    except Exception as cleanup_error:
                        log_func(f"✗ Failed to cleanup {item.name}: {cleanup_error}")
                
                raise Exception(f"Repository verification failed: {result.stderr}")
            
            log_func("✓ Repository verification successful!")
            
            # Parse snapshots to get statistics
            log_func("Parsing repository statistics...")
            try:
                snapshots = json.loads(result.stdout) if result.stdout.strip() else []
                snapshot_count = len(snapshots)
                latest_snapshot = snapshots[-1]['short_id'] if snapshots else None
                log_func(f"✓ Found {snapshot_count} snapshots")
                if latest_snapshot:
                    log_func(f"✓ Latest snapshot: {latest_snapshot}")
            except (json.JSONDecodeError, KeyError, IndexError) as e:
                log_func(f"⚠ Could not parse snapshots: {e}")
                snapshot_count = 0
                latest_snapshot = None
            
            # Calculate repository size
            log_func("Calculating repository size...")
            status_var.set("Calculating size...")
            repo_size_gb = self.calculate_repo_size(dest_repo)
            log_func(f"✓ Repository size: {repo_size_gb:.1f} GB")
            
            # Create database entry
            log_func("Creating database entry...")
            status_var.set("Creating database record...")
            image_id = generate_uuidv7()
            log_func(f"Generated image ID: {image_id}")
            
            self.db.create_image(
                image_id=image_id,
                client_id=client_id,
                site_id=site_id,
                role=role,
                repository_path=str(dest_repo),
                repository_size_gb=repo_size_gb,
                snapshot_count=snapshot_count,
                latest_snapshot_id=latest_snapshot,
                restic_password=password
            )
            
            log_func(f"✓ Repository registered with ID: {image_id}")
            log_func(f"✓ Repository organized under client: {client_id}")
            log_func(f"✓ Final path: {dest_repo}")
            log_func(f"✓ Snapshots: {snapshot_count}")
            log_func(f"✓ Size: {repo_size_gb:.1f} GB")
            
            # Create JSON metadata file in client repository folder
            log_func("Creating client metadata JSON file...")
            
            # Get client and site information from database
            clients = self.db.get_clients()
            sites = self.db.get_sites()
            
            client_info = None
            site_info = None
            
            for client in clients:
                if client[0] == client_id:  # client[0] is the ID
                    client_info = {
                        "id": client[0],
                        "name": client[1],
                        "short_name": client[2],
                        "description": client[3] if len(client) > 3 else ""
                    }
                    break
            
            for site in sites:
                if site[0] == site_id:  # site[0] is the ID
                    site_info = {
                        "id": site[0],
                        "name": site[2],  # site[2] is name
                        "short_name": site[3],  # site[3] is short_name
                        "description": site[4] if len(site) > 4 else ""
                    }
                    break
            
            image_info = {
                "id": image_id,
                "role": role,
                "site_id": site_id,
                "repository_path": str(dest_repo),
                "snapshot_count": snapshot_count,
                "latest_snapshot_id": latest_snapshot,
                "repository_size_gb": repo_size_gb
            }
            
            if self.create_client_metadata_json(client_id, client_info, site_info, image_info):
                log_func("✓ Client metadata JSON file created")
            else:
                log_func("⚠ Failed to create client metadata JSON file")
            
            status_var.set("Import completed successfully!")
            return True
            
        except Exception as e:
            log_func(f"✗ Repository import failed: {str(e)}")
            status_var.set("Import failed!")
            return False

    def perform_repository_import(self, source_repo, client_id, site_id, role, password):
        """Perform the actual repository import with client UUID directory structure"""
        try:
            source_path = Path(source_repo)
            
            # Create destination path: {restic_base}/{client_uuid} (no subfolder)
            restic_base = self.get_restic_base_path()
            dest_repo = restic_base / client_id
            
            self.log_step2(f"Destination path: {dest_repo}")
            
            # Check if destination already exists and has repository files
            if dest_repo.exists() and any(dest_repo.iterdir()):
                raise Exception(f"Client directory already contains files. Cannot import repository.")
            
            # Create client directory if needed
            dest_repo.mkdir(parents=True, exist_ok=True)
            self.log_step2(f"Created client directory: {dest_repo}")
            
            # Copy repository contents (not the folder itself)
            self.log_step2("Copying repository files...")
            for item in Path(source_repo).iterdir():
                if item.is_dir():
                    shutil.copytree(item, dest_repo / item.name)
                else:
                    shutil.copy2(item, dest_repo / item.name)
            self.log_step2("Repository files copied successfully")
            
            # Verify repository integrity
            self.log_step2("Verifying repository integrity...")
            restic_exe = self.download_restic()
            if not restic_exe:
                # Clean up on failure - remove all copied files from client directory
                for item in dest_repo.iterdir():
                    if item.is_dir():
                        shutil.rmtree(item)
                    else:
                        item.unlink()
                raise Exception("Could not obtain restic binary")
            
            # Test repository access
            os.environ['RESTIC_REPOSITORY'] = str(dest_repo)
            os.environ['RESTIC_PASSWORD'] = password
            
            check_cmd = [restic_exe, "snapshots", "--json"]
            result = subprocess.run(check_cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
            
            if result.returncode != 0:
                # Clean up on failure - remove all copied files from client directory
                for item in dest_repo.iterdir():
                    if item.is_dir():
                        shutil.rmtree(item)
                    else:
                        item.unlink()
                raise Exception(f"Repository verification failed: {result.stderr}")
            
            # Parse snapshots to get statistics
            try:
                snapshots = json.loads(result.stdout) if result.stdout.strip() else []
                snapshot_count = len(snapshots)
                latest_snapshot = snapshots[-1]['short_id'] if snapshots else None
            except (json.JSONDecodeError, KeyError, IndexError):
                snapshot_count = 0
                latest_snapshot = None
            
            # Calculate repository size
            repo_size_gb = self.calculate_repo_size(dest_repo)
            
            # Create database entry
            image_id = generate_uuidv7()
            self.db.create_image(
                image_id=image_id,
                client_id=client_id,
                site_id=site_id,
                role=role,
                repository_path=str(dest_repo),
                repository_size_gb=repo_size_gb,
                snapshot_count=snapshot_count,
                latest_snapshot_id=latest_snapshot,
                restic_password=password
            )
            
            # Create JSON metadata file in client repository folder
            self.log_step2("Creating client metadata JSON file...")
            
            # Get client and site information from database
            clients = self.db.get_clients()
            sites = self.db.get_sites()
            
            client_info = None
            site_info = None
            
            for client in clients:
                if client[0] == client_id:  # client[0] is the ID
                    client_info = {
                        "id": client[0],
                        "name": client[1],
                        "short_name": client[2],
                        "description": client[3] if len(client) > 3 else ""
                    }
                    break
            
            for site in sites:
                if site[0] == site_id:  # site[0] is the ID
                    site_info = {
                        "id": site[0],
                        "name": site[2],  # site[2] is name
                        "short_name": site[3],  # site[3] is short_name
                        "description": site[4] if len(site) > 4 else ""
                    }
                    break
            
            image_info = {
                "id": image_id,
                "role": role,
                "site_id": site_id,
                "repository_path": str(dest_repo),
                "snapshot_count": snapshot_count,
                "latest_snapshot_id": latest_snapshot,
                "repository_size_gb": repo_size_gb
            }
            
            if self.create_client_metadata_json(client_id, client_info, site_info, image_info):
                self.log_step2("Client metadata JSON file created successfully")
            else:
                self.log_step2("Warning: Failed to create client metadata JSON file")
            
            self.log_step2(f"Repository registered with ID: {image_id}")
            self.log_step2(f"Repository organized under client: {client_id}")
            self.log_step2(f"Final path: {dest_repo}")
            self.log_step2(f"Snapshots: {snapshot_count}")
            self.log_step2(f"Size: {repo_size_gb:.1f} GB")
            
            return True
            
        except Exception as e:
            self.log_step2(f"Repository import failed: {str(e)}")
            return False
            
    def update_repo_name(self):
        """Update repository name based on client selection"""
        if not hasattr(self, 'repo_name_var'):
            return  # UI not fully initialized yet
            
        client_name = self.client_var.get()
        if client_name and client_name != "-- Select Client --":
            # Clean client name for use in repository name
            clean_name = "".join(c for c in client_name if c.isalnum() or c in "-_").lower()
            repo_name = f"{clean_name}-restic-repo"
            self.repo_name_var.set(repo_name)

    def on_size_changed(self, value):
        """Handle VHDX size slider change"""
        size = int(float(value))
        self.size_label.config(text=f"{size} GB")
    
    def on_repo_type_changed(self):
        """Handle repository type selection change"""
        if not hasattr(self, 'repo_type_var'):
            return
            
        repo_type = self.repo_type_var.get()
        
        if repo_type == "local":
            # Hide S3 frame for local repos
            if hasattr(self, 's3_config_frame'):
                self.s3_config_frame.grid_remove()
        else:  # s3
            # Show S3 frame for S3 repos
            if hasattr(self, 's3_config_frame'):
                self.s3_config_frame.grid()
                self.update_s3_status()
    
    def update_s3_status(self):
        """Update S3 configuration status display"""
        if not hasattr(self, 's3_status_var'):
            return
            
        s3_config = self.db.get_s3_config()
        if s3_config:
            self.s3_status_var.set(f"✓ S3 configured: {s3_config.get('s3_bucket', 'Unknown bucket')}")
            if hasattr(self, 's3_status_label'):
                self.s3_status_label.config(foreground="green")
        else:
            self.s3_status_var.set("⚠ S3 not configured - click 'Configure S3...' button")
            if hasattr(self, 's3_status_label'):
                self.s3_status_label.config(foreground="red")

    def start_professional_image_creation(self):
        """Start the restic repository creation/import process"""
        # Validate inputs
        if not self.client_var.get():
            messagebox.showerror("Error", "Please select or create a client")
            return
        
        if not self.site_var.get():
            messagebox.showerror("Error", "Please select or create a site")
            return
        
        if self.import_existing_var.get():
            # Import existing repository
            if not self.import_repo_var.get() or not Path(self.import_repo_var.get()).exists():
                messagebox.showerror("Error", "Please select a valid existing repository")
                return
            action_text = "Import existing restic repository? This will:\n• Copy repository to managed location\n• Register in database\n• Scan for existing snapshots\n\nContinue?"
        else:
            # Create new repository
            if not self.repo_location_var.get():
                messagebox.showerror("Error", "Please select a repository location")
                return
            if not self.repo_password_var.get():
                messagebox.showerror("Error", "Please provide a repository password")
                return
            action_text = "Create new restic repository? This will:\n• Initialize new restic repository\n• Set up encryption with password\n• Register in database\n• Prepare for backups\n\nContinue?"
        
        if not messagebox.askyesno("Confirm", action_text):
            return
        
        self.create_image_button.config(state="disabled")
        thread = threading.Thread(target=self.repository_creation_worker)
        thread.daemon = True
        thread.start()

    def repository_creation_worker(self):
        """Worker thread for repository creation/import"""
        try:
            if self.import_existing_var.get():
                success = self.import_existing_repository()
            else:
                success = self.create_new_repository()
                
            if success:
                self.log("SUCCESS: Repository operation completed successfully!")
                # Refresh the repository browser
                if hasattr(self, 'refresh_images_list'):
                    self.refresh_images_list()
            else:
                self.log("ERROR: Repository operation failed!")
                
        except Exception as e:
            self.log(f"FATAL: Repository operation failed: {e}")
        finally:
            # Re-enable the create button
            self.create_image_button.config(state="normal")
            
    def create_new_repository(self):
        """Create a new restic repository"""
        try:
            # Get form data
            client_name = self.client_var.get()
            site_name = self.site_var.get()
            role = self.role_var.get()
            repo_location = self.repo_location_var.get()
            repo_password = self.repo_password_var.get()
            repo_name = self.repo_name_var.get()
            
            # Get client and site IDs
            clients = self.db.get_clients()
            client_id = None
            for cid, name, _, _ in clients:
                if name == client_name:
                    client_id = cid
                    break
                    
            sites = self.db.get_sites(client_id) if client_id else []
            site_id = None
            for sid, _, name, _, _, _ in sites:
                if name == site_name:
                    site_id = sid
                    break
            
            if not client_id or not site_id:
                self.log("ERROR: Could not find client or site IDs")
                return False
            
            # Create repository directory organized by client UUID
            # Use the configured restic base path instead of user-specified location
            restic_base = self.get_restic_base_path()
            client_repo_dir = restic_base / client_id  # Use client UUID as directory name
            repo_path = client_repo_dir / repo_name
            repo_path.mkdir(parents=True, exist_ok=True)
            
            self.log(f"INFO: Creating new restic repository at: {repo_path}")
            
            # Initialize restic repository
            restic_exe = self.download_restic()
            if not restic_exe:
                self.log("ERROR: Could not obtain restic binary")
                return False
            
            # Set environment variables
            os.environ['RESTIC_REPOSITORY'] = str(repo_path)
            os.environ['RESTIC_PASSWORD'] = repo_password
            
            # Initialize repository
            init_cmd = [restic_exe, "init"]
            result = subprocess.run(init_cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
            
            if result.returncode != 0:
                self.log(f"ERROR: Failed to initialize restic repository: {result.stderr}")
                return False
            
            self.log("SUCCESS: Restic repository initialized")
            
            # Calculate initial size (should be minimal)
            repo_size_gb = self.calculate_repo_size(repo_path)
            
            # Create database entry
            image_id = generate_uuidv7()
            self.db.create_image(
                image_id=image_id,
                client_id=client_id,
                site_id=site_id,
                role=role,
                repository_path=str(repo_path),
                repository_size_gb=repo_size_gb,
                snapshot_count=0,
                restic_password=repo_password
            )
            
            self.log(f"SUCCESS: Repository registered in database with ID: {image_id}")
            
            # Create JSON metadata file in client repository folder
            client_info = {
                "id": client_id,
                "name": client_name,
                "short_name": "",  # Will be filled from database if available
                "description": ""
            }
            site_info = {
                "id": site_id,
                "name": site_name,
                "short_name": "",  # Will be filled from database if available  
                "description": ""
            }
            image_info = {
                "id": image_id,
                "role": role,
                "site_id": site_id,
                "repository_path": str(repo_path),
                "snapshot_count": 0,
                "latest_snapshot_id": "",
                "repository_size_gb": repo_size_gb
            }
            
            # Client metadata JSON will be created when first backup is taken
            
            return True
            
        except Exception as e:
            self.log(f"ERROR: Failed to create repository: {e}")
            return False
            
    def import_existing_repository(self):
        """Import an existing restic repository"""
        try:
            # Get form data
            client_name = self.client_var.get()
            site_name = self.site_var.get()
            role = self.role_var.get()
            source_repo = self.import_repo_var.get()
            repo_name = self.repo_name_var.get()
            
            # Get destination from repo location or use default
            if self.repo_location_var.get():
                dest_base = Path(self.repo_location_var.get())
            else:
                # Use default location
                dest_base = Path(os.environ.get('PUBLIC', 'C:\\Users\\Public')) / "Documents" / "restic-repos"
                dest_base.mkdir(parents=True, exist_ok=True)
            
            dest_repo = dest_base / repo_name
            
            self.log(f"INFO: Importing repository from: {source_repo}")
            self.log(f"INFO: Copying to: {dest_repo}")
            
            # Copy repository
            if dest_repo.exists():
                self.log("WARNING: Destination repository already exists")
                if not messagebox.askyesno("Confirm", f"Repository {dest_repo} already exists. Overwrite?"):
                    return False
                shutil.rmtree(dest_repo)
            
            shutil.copytree(source_repo, dest_repo)
            self.log("SUCCESS: Repository copied successfully")
            
            # Try to determine repository password (user will need to provide)
            repo_password = simpledialog.askstring("Repository Password", 
                                                "Please enter the password for the imported repository:",
                                                show='*')
            if not repo_password:
                self.log("ERROR: Repository password is required")
                return False
            
            # Verify repository is accessible
            restic_exe = self.download_restic()
            if not restic_exe:
                self.log("ERROR: Could not obtain restic binary")
                return False
            
            os.environ['RESTIC_REPOSITORY'] = str(dest_repo)
            os.environ['RESTIC_PASSWORD'] = repo_password
            
            # List snapshots to verify repository
            list_cmd = [restic_exe, "snapshots", "--json"]
            result = subprocess.run(list_cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
            
            if result.returncode != 0:
                self.log(f"ERROR: Could not access repository with provided password: {result.stderr}")
                return False
            
            # Parse snapshot information
            try:
                snapshots = json.loads(result.stdout) if result.stdout.strip() else []
                snapshot_count = len(snapshots)
                latest_snapshot_id = snapshots[-1]['id'][:8] if snapshots else None
                self.log(f"INFO: Found {snapshot_count} snapshots in repository")
            except:
                snapshot_count = 0
                latest_snapshot_id = None
                
            # Get client and site IDs
            clients = self.db.get_clients()
            client_id = None
            for cid, name, _, _ in clients:
                if name == client_name:
                    client_id = cid
                    break
                    
            sites = self.db.get_sites(client_id) if client_id else []
            site_id = None
            for sid, _, name, _, _, _ in sites:
                if name == site_name:
                    site_id = sid
                    break
            
            if not client_id or not site_id:
                self.log("ERROR: Could not find client or site IDs")
                return False
            
            # Calculate repository size
            repo_size_gb = self.calculate_repo_size(dest_repo)
            
            # Create database entry
            image_id = generate_uuidv7()
            self.db.create_image(
                image_id=image_id,
                client_id=client_id,
                site_id=site_id,
                role=role,
                repository_path=str(dest_repo),
                repository_size_gb=repo_size_gb,
                snapshot_count=snapshot_count,
                latest_snapshot_id=latest_snapshot_id,
                restic_password=repo_password
            )
            
            self.log(f"SUCCESS: Repository imported and registered with ID: {image_id}")
            return True
            
        except Exception as e:
            self.log(f"ERROR: Failed to import repository: {e}")
            return False
            
    def calculate_repo_size(self, repo_path):
        """Calculate repository size in GB"""
        try:
            total_size = 0
            for dirpath, dirnames, filenames in os.walk(repo_path):
                for filename in filenames:
                    filepath = Path(dirpath) / filename
                    if filepath.exists():
                        total_size += filepath.stat().st_size
            return max(1, round(total_size / (1024**3)))  # Convert to GB, minimum 1GB
        except:
            return 1

    def restore_repository_to_vhdx(self, repository_path, restic_password, snapshot_id=None, vhdx_size_gb=256):
        """Restore a restic repository snapshot to a working VHDX file"""
        try:
            # Get working directory
            working_dir = self.db.get_working_vhdx_directory()
            
            # Generate unique VHDX filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            vhdx_filename = f"working_image_{timestamp}.vhdx"
            vhdx_path = working_dir / vhdx_filename
            
            self.log(f"INFO: Starting restore to working VHDX: {vhdx_path}")
            
            # Step 1: Create new VHDX file
            self.log("INFO: Step 1 - Creating VHDX file...")
            if not self.create_vhdx_file(vhdx_path, vhdx_size_gb):
                return None
                
            # Step 2: Initialize VHDX with GPT partitioning
            self.log("INFO: Step 2 - Initializing VHDX with GPT partitioning...")
            if not self.initialize_vhdx_gpt(vhdx_path):
                return None
            
            # Step 3: Mount VHDX
            self.log("INFO: Step 3 - Mounting VHDX...")
            mount_point = self.mount_vhdx(vhdx_path)
            if not mount_point:
                return None
                
            try:
                # Step 4: Restore from restic repository
                self.log("INFO: Step 4 - Restoring from restic repository...")
                if not self.restore_restic_to_mount(repository_path, restic_password, mount_point, snapshot_id):
                    return None
                    
                self.log(f"SUCCESS: Repository restored to working VHDX: {vhdx_path}")
                return vhdx_path
                
            finally:
                # Always unmount VHDX
                self.log("INFO: Unmounting VHDX...")
                self.unmount_vhdx(vhdx_path)
                
        except Exception as e:
            self.log(f"ERROR: Failed to restore repository to VHDX: {e}")
            return None
            
    def create_vhdx_file(self, vhdx_path, size_gb):
        """Create a new VHDX file using diskpart"""
        try:
            diskpart_script = f'''
create vdisk file="{vhdx_path}" maximum={size_gb * 1024} type=expandable
'''
            script_path = vhdx_path.parent / f"create_vhdx_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            script_path.write_text(diskpart_script, encoding='utf-8')
            
            cmd = ["diskpart", "/s", str(script_path)]
            result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
            
            # Clean up script file
            try:
                script_path.unlink()
            except:
                pass
                
            if result.returncode == 0:
                self.log(f"SUCCESS: Created VHDX file: {vhdx_path}")
                return True
            else:
                self.log(f"ERROR: Failed to create VHDX: {result.stderr}")
                return False
                
        except Exception as e:
            self.log(f"ERROR: Exception creating VHDX: {e}")
            return False
            
    def initialize_vhdx_gpt(self, vhdx_path):
        """Initialize VHDX with GPT partitioning scheme"""
        try:
            diskpart_script = f'''
select vdisk file="{vhdx_path}"
attach vdisk
convert gpt
create partition efi size=100
assign letter=s
format quick fs=fat32 label="System"
create partition msr size=16
create partition primary
assign letter=w
format quick fs=ntfs label="Windows"
active
detach vdisk
'''
            script_path = vhdx_path.parent / f"init_vhdx_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            script_path.write_text(diskpart_script, encoding='utf-8')
            
            cmd = ["diskpart", "/s", str(script_path)]
            result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
            
            # Clean up script file
            try:
                script_path.unlink()
            except:
                pass
                
            if result.returncode == 0:
                self.log("SUCCESS: VHDX initialized with GPT partitioning")
                return True
            else:
                self.log(f"ERROR: Failed to initialize VHDX: {result.stderr}")
                return False
                
        except Exception as e:
            self.log(f"ERROR: Exception initializing VHDX: {e}")
            return False
            
    def mount_vhdx(self, vhdx_path):
        """Mount VHDX and return the Windows partition mount point"""
        try:
            diskpart_script = f'''
select vdisk file="{vhdx_path}"
attach vdisk
list partition
'''
            script_path = vhdx_path.parent / f"mount_vhdx_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            script_path.write_text(diskpart_script, encoding='utf-8')
            
            cmd = ["diskpart", "/s", str(script_path)]
            result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
            
            # Clean up script file
            try:
                script_path.unlink()
            except:
                pass
                
            if result.returncode == 0:
                # Try to find the mounted Windows partition
                # Look for the drive that was assigned letter 'W'
                mount_point = "W:\\"
                if os.path.exists(mount_point):
                    self.log(f"SUCCESS: VHDX mounted at {mount_point}")
                    return mount_point
                else:
                    self.log("ERROR: Could not find mounted Windows partition")
                    return None
            else:
                self.log(f"ERROR: Failed to mount VHDX: {result.stderr}")
                return None
                
        except Exception as e:
            self.log(f"ERROR: Exception mounting VHDX: {e}")
            return None
            
    def unmount_vhdx(self, vhdx_path):
        """Unmount VHDX file"""
        try:
            diskpart_script = f'''
select vdisk file="{vhdx_path}"
detach vdisk
'''
            script_path = vhdx_path.parent / f"unmount_vhdx_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            script_path.write_text(diskpart_script, encoding='utf-8')
            
            cmd = ["diskpart", "/s", str(script_path)]
            result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
            
            # Clean up script file
            try:
                script_path.unlink()
            except:
                pass
                
            if result.returncode == 0:
                self.log("SUCCESS: VHDX unmounted")
                return True
            else:
                self.log(f"WARNING: Failed to unmount VHDX: {result.stderr}")
                return False
                
        except Exception as e:
            self.log(f"WARNING: Exception unmounting VHDX: {e}")
            return False
            
    def restore_restic_to_mount(self, repository_path, restic_password, mount_point, snapshot_id=None):
        """Restore restic repository to mounted drive"""
        try:
            # Get restic binary
            restic_exe = self.download_restic()
            if not restic_exe:
                self.log("ERROR: Could not obtain restic binary")
                return False
            
            # Set environment variables
            os.environ['RESTIC_REPOSITORY'] = str(repository_path)
            os.environ['RESTIC_PASSWORD'] = restic_password
            
            # Determine which snapshot to restore
            if not snapshot_id:
                # Get latest snapshot
                list_cmd = [restic_exe, "snapshots", "--json"]
                result = subprocess.run(list_cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
                
                if result.returncode != 0:
                    self.log(f"ERROR: Could not list snapshots: {result.stderr}")
                    return False
                
                try:
                    snapshots = json.loads(result.stdout) if result.stdout.strip() else []
                    if not snapshots:
                        self.log("ERROR: No snapshots found in repository")
                        return False
                    snapshot_id = snapshots[-1]['id']
                    self.log(f"INFO: Using latest snapshot: {snapshot_id[:8]}")
                except Exception as e:
                    self.log(f"ERROR: Could not parse snapshots: {e}")
                    return False
            
            # Restore snapshot to mount point
            restore_cmd = [
                restic_exe, "restore", snapshot_id,
                "--target", mount_point,
                "--verbose"
            ]
            
            self.log(f"COMMAND: {' '.join(restore_cmd)}")
            self.log("INFO: Starting restic restore - this may take 10-30 minutes...")
            
            restore_proc = subprocess.Popen(
                restore_cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding='utf-8',
                errors='ignore'
            )
            
            # Stream restore output
            if restore_proc.stdout:
                for line in iter(restore_proc.stdout.readline, ''):
                    line = line.strip()
                    if line:
                        self.log(line)
            
            restore_proc.wait()
            
            if restore_proc.returncode == 0:
                self.log("SUCCESS: Restic restore completed!")
                return True
            else:
                self.log(f"ERROR: Restic restore failed with exit code: {restore_proc.returncode}")
                return False
                
        except Exception as e:
            self.log(f"ERROR: Exception during restic restore: {e}")
            return False

    def professional_image_creation_worker(self):
        """Worker thread for professional image creation"""
        try:
            self.log("=== STARTING PROFESSIONAL IMAGE CREATION ===")
            
            # Get parameters
            client_name = self.client_var.get()
            site_name = self.site_var.get()
            role = self.role_var.get()
            wim_source = self.wim_source_var.get()
            vhdx_size = self.vhdx_size_var.get()
            
            # Find client and site IDs
            clients = self.db.get_clients()
            client_id = None
            for cid, name, short_name, desc in clients:
                if name == client_name:
                    client_id = cid
                    break
            
            if not client_id:
                self.log("ERROR: Client not found in database")
                return
            
            sites = self.db.get_sites(client_id)
            site_id = None
            for sid, _, name, _, _, _ in sites:
                if name == site_name:
                    site_id = sid
                    break
            
            if not site_id:
                self.log("ERROR: Site not found in database")
                return
            
            # Generate image ID and paths
            image_id = generate_uuidv7()
            vhdx_filename = f"{image_id}.vhdx"
            vhdx_path = self.image_store_path / vhdx_filename
            
            self.log(f"INFO: Creating image {image_id}")
            self.log(f"INFO: Client: {client_name}, Site: {site_name}, Role: {role}")
            self.log(f"INFO: VHDX Size: {vhdx_size} GB")
            self.log(f"INFO: Source WIM: {wim_source}")
            self.log(f"INFO: VHDX Path: {vhdx_path}")
            
            # TODO: Implement the actual image creation process
            # This would include:
            # 1. Create blank VHDX file
            # 2. Mount and partition (GPT with EFI/Recovery/OS)
            # 3. Deploy WIM to OS partition
            # 4. Configure boot
            # 5. Create VM if requested
            # 6. Update database
            
            # For now, just placeholder
            self.log("INFO: Professional image creation is not yet fully implemented")
            self.log("INFO: This would create a complete deployment-ready image")
            
            # Simulate work
            import time
            time.sleep(3)
            
            self.log("SUCCESS: Professional image creation completed (placeholder)")
            
        except Exception as e:
            self.log(f"FATAL: Professional image creation failed: {e}")
        finally:
            self.create_image_button.config(state="normal")

    def refresh_client_site_data(self):
        """Refresh client and site data in UI"""
        try:
            # Refresh clients combo
            clients = self.db.get_clients()
            client_names = [name for _, name, _, _ in clients]
            
            # Update client combo values
            if hasattr(self, 'client_combo'):
                self.client_combo['values'] = tuple(client_names)
                # Auto-select first client if none selected and clients available
                if client_names and not self.client_var.get():
                    self.client_var.set(client_names[0])
                
                # Always trigger site population if a client is selected
                if self.client_var.get():
                    self.on_client_selected()
            
            # Clear site combo only if no client is selected
            elif hasattr(self, 'site_combo'):
                self.site_combo['values'] = ()
                self.site_var.set("")
            
            # Refresh database tab trees
            if hasattr(self, 'clients_tree'):
                self.clients_tree.delete(*self.clients_tree.get_children())
                for _, name, short_name, _ in clients:
                    self.clients_tree.insert("", "end", values=(name, short_name))
            
            if hasattr(self, 'sites_tree'):
                self.sites_tree.delete(*self.sites_tree.get_children())
                sites = self.db.get_sites()
                for _, _, name, short_name, _, client_name in sites:
                    self.sites_tree.insert("", "end", values=(client_name, name, short_name))
            
        except Exception as e:
            self.log(f"ERROR: Failed to refresh client/site data: {e}")

    def refresh_images_list(self):
        """Refresh the repositories list in browse tab"""
        try:
            self.images_tree.delete(*self.images_tree.get_children())
            images = self.db.get_images()
            
            for image_data in images:
                # Handle both legacy and new format
                if len(image_data) >= 17:  # New format with repository fields
                    (image_id, role, image_path, image_size_gb, vm_name, vm_created, 
                     status, created_at, client_name, client_short, site_name, site_short,
                     repository_path, repository_size_gb, snapshot_count, latest_snapshot_id, image_type) = image_data
                else:  # Legacy format
                    (image_id, role, image_path, image_size_gb, vm_name, vm_created, 
                     status, created_at, client_name, client_short, site_name, site_short) = image_data
                    repository_path = repository_size_gb = snapshot_count = latest_snapshot_id = image_type = None
                
                # Determine which data to show based on image type
                if image_type == 'restic' and repository_path:
                    size_display = f"{repository_size_gb} GB" if repository_size_gb else "Unknown"
                    snapshots_display = str(snapshot_count) if snapshot_count is not None else "0"
                    type_display = "Restic"
                else:
                    # Legacy VHDX format
                    size_display = f"{image_size_gb} GB" if image_size_gb else "Unknown"
                    snapshots_display = "N/A"
                    type_display = "Legacy"
                
                created_date = created_at.split()[0] if created_at else "Unknown"
                
                self.images_tree.insert("", "end", values=(
                    client_name, site_name, role, size_display, 
                    snapshots_display, type_display, created_date, status
                ), tags=(image_id,))
                
        except Exception as e:
            self.log(f"ERROR: Failed to refresh repositories list: {e}")

    def check_orphan_files(self):
        """Check for orphaned VHDX files not in database"""
        try:
            self.log("INFO: Checking for orphaned VHDX files...")
            
            # Get all image files in storage directory
            image_files = list(self.image_store_path.glob("*.wim")) + list(self.image_store_path.glob("*.vhdx"))
            
            # Get all known image IDs from database
            images = self.db.get_images()
            known_ids = set()
            for image_data in images:
                image_path = image_data[2]  # image_path column
                if image_path:
                    known_ids.add(Path(image_path).name)
            
            # Find orphans
            orphans = []
            for image_file in image_files:
                if image_file.name not in known_ids:
                    # Check if it has a corresponding metadata file
                    metadata_file = image_file.with_suffix('.metadata.json')
                    orphans.append((image_file, metadata_file.exists()))
            
            if orphans:
                self.log(f"INFO: Found {len(orphans)} orphaned image files:")
                for image_file, has_metadata in orphans:
                    metadata_status = "with metadata" if has_metadata else "no metadata"
                    self.log(f"  • {image_file.name} ({metadata_status})")
                    
                messagebox.showinfo("Orphan Check", 
                    f"Found {len(orphans)} orphaned image files.\n"
                    "Check the log for details.\n"
                    "Use 'Import Orphan' to add them to the database.")
            else:
                self.log("INFO: No orphaned image files found")
                messagebox.showinfo("Orphan Check", "No orphaned files found.")
                
        except Exception as e:
            self.log(f"ERROR: Orphan check failed: {e}")

    def import_orphan_file(self):
        """Import an orphaned image file into the database"""
        try:
            # Select orphaned file
            image_file = filedialog.askopenfilename(
                title="Select Orphaned Image File",
                initialdir=self.image_store_path,
                filetypes=[("Image Files", "*.wim;*.vhdx"), ("WIM Files", "*.wim"), ("VHDX Files", "*.vhdx")]
            )
            
            if not image_file:
                return
            
            image_path = Path(image_file)
            
            # Check if it has metadata
            metadata_path = image_path.with_suffix('.metadata.json')
            
            if metadata_path.exists():
                # Import from metadata
                self.import_from_metadata(metadata_path)
            else:
                # Manual import dialog
                self.manual_import_orphan(image_path)
                
        except Exception as e:
            self.log(f"ERROR: Failed to import orphan: {e}")

    def import_from_metadata(self, metadata_path):
        """Import orphan using existing metadata file"""
        try:
            with open(metadata_path, 'r') as f:
                metadata = json.load(f)
            
            # Validate metadata has required fields
            required_fields = ['client_name', 'site_name', 'role', 'vhdx_size_gb']
            for field in required_fields:
                if field not in metadata:
                    raise ValueError(f"Metadata missing required field: {field}")
            
            # Find or create client
            clients = self.db.get_clients()
            client_id = None
            for cid, name, _, _ in clients:
                if name == metadata['client_name']:
                    client_id = cid
                    break
            
            if not client_id:
                self.log(f"INFO: Creating new client: {metadata['client_name']}")
                client_id = self.db.add_client(metadata['client_name'], 
                                             metadata.get('client_short', metadata['client_name'][:8].upper()))
            
            # Find or create site
            sites = self.db.get_sites(client_id)
            site_id = None
            for sid, _, name, _, _, _ in sites:
                if name == metadata['site_name']:
                    site_id = sid
                    break
            
            if not site_id:
                self.log(f"INFO: Creating new site: {metadata['site_name']}")
                site_id = self.db.add_site(client_id, metadata['site_name'], 
                                         metadata.get('site_short', metadata['site_name'][:8].upper()))
            
            # Import image record
            vhdx_path = Path(metadata_path).with_suffix('.vhdx')
            image_id = self.db.add_image(
                client_id, site_id, metadata['role'], 
                metadata.get('wim_source_path', ''), 
                str(vhdx_path), metadata['vhdx_size_gb'],
                metadata.get('vm_name', '')
            )
            
            self.log(f"SUCCESS: Imported orphan image {image_id} from metadata")
            self.refresh_images_list()
            self.refresh_client_site_data()
            
        except Exception as e:
            self.log(f"ERROR: Failed to import from metadata: {e}")

    def manual_import_orphan(self, vhdx_path):
        """Manual import dialog for orphan without metadata"""
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Import Orphan: {vhdx_path.name}")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        ttk.Label(dialog, text=f"Import: {vhdx_path.name}", font=("TkDefaultFont", 12, "bold")).pack(pady=10)
        
        # Client selection
        ttk.Label(dialog, text="Client:").pack(anchor="w", padx=20)
        client_var = tk.StringVar()
        client_combo = ttk.Combobox(dialog, textvariable=client_var, width=40)
        client_combo.pack(padx=20, pady=5)
        
        # Site selection
        ttk.Label(dialog, text="Site:").pack(anchor="w", padx=20)
        site_var = tk.StringVar()
        site_combo = ttk.Combobox(dialog, textvariable=site_var, width=40)
        site_combo.pack(padx=20, pady=5)
        
        # Role
        ttk.Label(dialog, text="Role:").pack(anchor="w", padx=20)
        role_var = tk.StringVar()
        role_combo = ttk.Combobox(dialog, textvariable=role_var, width=40, values=[
            "Desktop", "Server", "Workstation", "Domain Controller", "Database", 
            "Web Server", "File Server", "Terminal Server", "Unknown"
        ])
        role_combo.pack(padx=20, pady=5)
        role_combo.set("Unknown")
        
        # Populate client combo
        clients = self.db.get_clients()
        client_names = [name for _, name, _, _ in clients]
        client_combo['values'] = client_names
        
        def on_client_change(event=None):
            client_name = client_var.get()
            if client_name:
                client_id = None
                for cid, name, _, _ in clients:
                    if name == client_name:
                        client_id = cid
                        break
                if client_id:
                    sites = self.db.get_sites(client_id)
                    site_names = [name for _, _, name, _, _, _ in sites]
                    site_combo['values'] = site_names
        
        client_combo.bind('<<ComboboxSelected>>', on_client_change)
        
        def save_import():
            if not client_var.get() or not site_var.get() or not role_var.get():
                messagebox.showerror("Error", "Please fill all fields")
                return
            
            try:
                # Find client and site IDs
                client_id = None
                for cid, name, _, _ in clients:
                    if name == client_var.get():
                        client_id = cid
                        break
                
                if not client_id:
                    messagebox.showerror("Error", "Client not found")
                    return
                
                sites = self.db.get_sites(client_id)
                site_id = None
                for sid, _, name, _, _, _ in sites:
                    if name == site_var.get():
                        site_id = sid
                        break
                
                if not site_id:
                    messagebox.showerror("Error", "Site not found")
                    return
                
                # Get file size
                vhdx_size_gb = int(vhdx_path.stat().st_size / (1024**3))
                
                # Import
                image_id = self.db.add_image(
                    client_id, site_id, role_var.get(), 
                    '', str(vhdx_path), vhdx_size_gb
                )
                
                self.log(f"SUCCESS: Manually imported orphan image {image_id}")
                self.refresh_images_list()
                dialog.destroy()
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to import: {e}")
        
        ttk.Button(dialog, text="Import", command=save_import).pack(pady=20)

    def show_image_context_menu(self, event):
        """Show context menu for images"""
        selection = self.images_tree.selection()
        if not selection:
            return
        
        # Create context menu
        context_menu = tk.Menu(self.root, tearoff=0)
        context_menu.add_command(label="📋 View Details", command=lambda: self.show_image_details(selection[0]))
        context_menu.add_separator()
        context_menu.add_command(label="💾 Restore to Working VHDX", command=lambda: self.restore_image_to_vhdx(selection[0]))
        context_menu.add_command(label="📊 Browse Snapshots", command=lambda: self.browse_snapshots(selection[0]))
        context_menu.add_separator()
        context_menu.add_command(label="🗑️ Delete Repository", command=lambda: self.delete_image_repository(selection[0]))
        
        try:
            context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            context_menu.grab_release()

    def on_image_double_click(self, event):
        """Handle double-click on image"""
        selection = self.images_tree.selection()
        if not selection:
            return
        
        # Double-click triggers restore to working VHDX
        self.restore_image_to_vhdx(selection[0])

    def show_repository_details(self, item_id):
        """Show detailed information about a repository"""
        try:
            item = self.images_tree.item(item_id)
            values = item['values']
            if not values:
                return
            
            # Get repository details from database
            images = self.db.get_images()
            repo_data = None
            for image in images:
                client_name = self.db.get_client_name(image[1])
                site_name = self.db.get_site_name(image[2])
                if (client_name == values[0] and site_name == values[1] and 
                    image[3] == values[2]):  # role matches
                    repo_data = image
                    break
            
            if not repo_data:
                messagebox.showerror("Error", "Repository not found in database")
                return
            
            # Create details window
            details_window = tk.Toplevel(self.root)
            details_window.title(f"Repository Details - {values[0]}/{values[1]}/{values[2]}")
            details_window.geometry("600x500")
            details_window.transient(self.root)
            
            # Repository information
            info_frame = ttk.LabelFrame(details_window, text="Repository Information", padding="10")
            info_frame.pack(fill="x", padx=10, pady=10)
            
            info_text = f"""Client: {values[0]}
Site: {values[1]}
Role: {values[2]}
Repository Path: {repo_data[4]}
Size: {values[3]}
Snapshot Count: {values[4]}
Latest Snapshot: {repo_data[7] or 'None'}
Status: {values[7]}
Created: {values[6]}
Updated: {repo_data[9]}"""
            
            ttk.Label(info_frame, text=info_text, justify="left").pack(anchor="w")
            
            # Actions frame
            actions_frame = ttk.LabelFrame(details_window, text="Actions", padding="10")
            actions_frame.pack(fill="x", padx=10, pady=10)
            
            ttk.Button(actions_frame, text="💾 Restore to Working VHDX", 
                      command=lambda: [details_window.destroy(), self.restore_image_to_vhdx(item_id)]).pack(side="left", padx=5)
            ttk.Button(actions_frame, text="📊 Browse Snapshots", 
                      command=lambda: [details_window.destroy(), self.browse_snapshots(item_id)]).pack(side="left", padx=5)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to show details: {e}")

    def restore_image_to_vhdx(self, item_id):
        """Restore a repository to a working VHDX file"""
        try:
            item = self.images_tree.item(item_id)
            values = item['values']
            if not values:
                return
            
            # Get repository details from database
            images = self.db.get_images()
            repo_data = None
            for image in images:
                client_name = self.db.get_client_name(image[1])
                site_name = self.db.get_site_name(image[2])
                if (client_name == values[0] and site_name == values[1] and 
                    image[3] == values[2]):  # role matches
                    repo_data = image
                    break
            
            if not repo_data:
                messagebox.showerror("Error", "Repository not found in database")
                return
            
            # Create restore dialog
            restore_window = tk.Toplevel(self.root)
            restore_window.title(f"Restore {values[0]}/{values[1]}/{values[2]} to Working VHDX")
            restore_window.geometry("500x400")
            restore_window.transient(self.root)
            restore_window.grab_set()
            
            # Repository info
            info_frame = ttk.LabelFrame(restore_window, text="Repository", padding="10")
            info_frame.pack(fill="x", padx=10, pady=10)
            
            ttk.Label(info_frame, text=f"Client: {values[0]}", font=("TkDefaultFont", 9, "bold")).pack(anchor="w")
            ttk.Label(info_frame, text=f"Site: {values[1]}", font=("TkDefaultFont", 9, "bold")).pack(anchor="w")
            ttk.Label(info_frame, text=f"Role: {values[2]}", font=("TkDefaultFont", 9, "bold")).pack(anchor="w")
            ttk.Label(info_frame, text=f"Snapshots Available: {values[4]}").pack(anchor="w")
            
            # VHDX options
            options_frame = ttk.LabelFrame(restore_window, text="VHDX Options", padding="10")
            options_frame.pack(fill="x", padx=10, pady=10)
            
            ttk.Label(options_frame, text="VHDX Size (GB):").pack(anchor="w")
            size_var = tk.IntVar(value=256)
            size_spinbox = ttk.Spinbox(options_frame, from_=64, to=2048, textvariable=size_var, width=10)
            size_spinbox.pack(anchor="w", pady=5)
            
            # Snapshot selection
            snapshot_frame = ttk.LabelFrame(restore_window, text="Snapshot Selection", padding="10")
            snapshot_frame.pack(fill="both", expand=True, padx=10, pady=10)
            
            ttk.Label(snapshot_frame, text="Select snapshot (leave empty for latest):").pack(anchor="w")
            snapshot_var = tk.StringVar()
            snapshot_entry = ttk.Entry(snapshot_frame, textvariable=snapshot_var, width=50)
            snapshot_entry.pack(fill="x", pady=5)
            
            ttk.Label(snapshot_frame, text="💡 Use 'latest' or leave empty for most recent snapshot", 
                     font=("TkDefaultFont", 8), foreground="gray").pack(anchor="w")
            
            # Progress frame
            progress_frame = ttk.Frame(restore_window)
            progress_frame.pack(fill="x", padx=10, pady=10)
            
            progress_var = tk.StringVar(value="Ready to restore...")
            ttk.Label(progress_frame, textvariable=progress_var).pack(anchor="w")
            
            progress_bar = ttk.Progressbar(progress_frame, mode='indeterminate')
            progress_bar.pack(fill="x", pady=5)
            
            # Buttons
            button_frame = ttk.Frame(restore_window)
            button_frame.pack(fill="x", padx=10, pady=10)
            
            def start_restore():
                """Start the restore process in a thread"""
                size_gb = size_var.get()
                snapshot_id = snapshot_var.get().strip() or None
                
                if messagebox.askyesno("Confirm Restore", 
                                     f"Restore {values[0]}/{values[1]}/{values[2]} to new working VHDX?\n\n"
                                     f"Size: {size_gb} GB\n"
                                     f"Snapshot: {snapshot_id or 'latest'}\n\n"
                                     "This may take several minutes."):
                    
                    progress_bar.start()
                    progress_var.set("Starting restore...")
                    
                    def restore_thread():
                        try:
                            self.log_step2(f"Starting restore of {values[0]}/{values[1]}/{values[2]} to VHDX...")
                            progress_var.set("Restoring repository to VHDX...")
                            
                            # Call the restore method
                            vhdx_path = self.restore_repository_to_vhdx(
                                repo_data[4],  # repository_path
                                repo_data[8],  # restic_password
                                snapshot_id,
                                size_gb
                            )
                            
                            self.log_step2(f"Successfully restored to: {vhdx_path}")
                            
                            # Update UI on completion
                            restore_window.after(0, lambda: restore_complete(vhdx_path))
                            
                        except Exception as e:
                            self.log_step2(f"Restore failed: {str(e)}")
                            restore_window.after(0, lambda: restore_failed(str(e)))
                    
                    def restore_complete(vhdx_path):
                        progress_bar.stop()
                        progress_var.set("Restore completed!")
                        messagebox.showinfo("Restore Complete", 
                                          f"Repository restored to working VHDX:\n{vhdx_path}\n\n"
                                          "The VHDX is now mounted and ready for use.")
                        restore_window.destroy()
                    
                    def restore_failed(error):
                        progress_bar.stop()
                        progress_var.set("Restore failed!")
                        messagebox.showerror("Restore Failed", f"Failed to restore repository:\n{error}")
                    
                    threading.Thread(target=restore_thread, daemon=True).start()
            
            ttk.Button(button_frame, text="🚀 Start Restore", command=start_restore).pack(side="left", padx=5)
            ttk.Button(button_frame, text="Cancel", command=restore_window.destroy).pack(side="left", padx=5)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to start restore: {e}")

    def browse_snapshots(self, item_id):
        """Browse snapshots in a repository"""
        try:
            item = self.images_tree.item(item_id)
            values = item['values']
            if not values:
                return
            
            # Get repository details from database
            images = self.db.get_images()
            repo_data = None
            for image in images:
                client_name = self.db.get_client_name(image[1])
                site_name = self.db.get_site_name(image[2])
                if (client_name == values[0] and site_name == values[1] and 
                    image[3] == values[2]):  # role matches
                    repo_data = image
                    break
            
            if not repo_data:
                messagebox.showerror("Error", "Repository not found in database")
                return
            
            # Create snapshots browser window
            snapshots_window = tk.Toplevel(self.root)
            snapshots_window.title(f"Snapshots - {values[0]}/{values[1]}/{values[2]}")
            snapshots_window.geometry("800x600")
            snapshots_window.transient(self.root)
            
            # Repository info
            info_frame = ttk.LabelFrame(snapshots_window, text="Repository", padding="10")
            info_frame.pack(fill="x", padx=10, pady=10)
            
            ttk.Label(info_frame, text=f"Repository: {values[0]}/{values[1]}/{values[2]}").pack(anchor="w")
            ttk.Label(info_frame, text=f"Path: {repo_data[4]}").pack(anchor="w")
            
            # Snapshots list
            list_frame = ttk.LabelFrame(snapshots_window, text="Available Snapshots", padding="10")
            list_frame.pack(fill="both", expand=True, padx=10, pady=10)
            
            # Note: For now, just show placeholder - full implementation would require calling restic snapshots
            ttk.Label(list_frame, text="Snapshot browsing functionality will be implemented in future updates.\n\n"
                                    "For now, use the restore dialog and leave snapshot field empty for latest,\n"
                                    "or specify a snapshot ID if you know it.").pack(pady=20)
            
            ttk.Button(list_frame, text="Close", command=snapshots_window.destroy).pack(pady=10)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to browse snapshots: {e}")

    def delete_image_repository(self, item_id):
        """Delete a repository from the database (and optionally from disk)"""
        try:
            item = self.images_tree.item(item_id)
            values = item['values']
            if not values:
                return
            
            # Get repository details from database
            images = self.db.get_images()
            repo_data = None
            for image in images:
                client_name = self.db.get_client_name(image[1])
                site_name = self.db.get_site_name(image[2])
                if (client_name == values[0] and site_name == values[1] and 
                    image[3] == values[2]):  # role matches
                    repo_data = image
                    break
            
            if not repo_data:
                messagebox.showerror("Error", "Repository not found in database")
                return
            
            # Confirmation dialog
            if messagebox.askyesno("Confirm Deletion", 
                                 f"Delete repository {values[0]}/{values[1]}/{values[2]}?\n\n"
                                 f"This will remove the repository from the database.\n"
                                 f"The actual repository files will remain on disk.\n\n"
                                 "This action cannot be undone."):
                
                # Delete from database
                with sqlite3.connect(self.db.db_path) as conn:
                    cursor = conn.cursor()
                    cursor.execute("DELETE FROM images WHERE id = ?", (repo_data[0],))
                    conn.commit()
                
                self.log_step2(f"Deleted repository {values[0]}/{values[1]}/{values[2]} from database")
                messagebox.showinfo("Deleted", "Repository removed from database")
                
                # Refresh the images list
                self.refresh_images_list()
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete repository: {e}")

    def restore_selected_repository(self):
        """Restore the currently selected repository"""
        selection = self.images_tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a repository to restore")
            return
        
        self.restore_image_to_vhdx(selection[0])

    def import_repository_dialog(self):
        """Show dialog to import an existing repository"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Import Existing Repository")
        dialog.geometry("700x600")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # Title
        ttk.Label(dialog, text="Import Existing Restic Repository", 
                 font=("TkDefaultFont", 14, "bold")).pack(pady=20)
        
        # Source repository selection
        source_frame = ttk.LabelFrame(dialog, text="Source Repository", padding="10")
        source_frame.pack(fill="x", padx=20, pady=10)
        
        ttk.Label(source_frame, text="Select the restic repository to import:").pack(anchor="w")
        
        source_var = tk.StringVar()
        source_entry = ttk.Entry(source_frame, textvariable=source_var, width=70)
        source_entry.pack(fill="x", pady=5)
        
        def browse_source():
            path = filedialog.askdirectory(title="Select Restic Repository to Import")
            if path:
                source_var.set(path)
        
        ttk.Button(source_frame, text="Browse...", command=browse_source).pack(pady=5)
        
        # Client/Site selection
        client_site_frame = ttk.LabelFrame(dialog, text="Client & Site Assignment", padding="10")
        client_site_frame.pack(fill="x", padx=20, pady=10)
        
        # Client dropdown
        ttk.Label(client_site_frame, text="Assign to Client:").pack(anchor="w")
        client_var = tk.StringVar()
        client_combo = ttk.Combobox(client_site_frame, textvariable=client_var, width=50, state="readonly")
        client_combo.pack(fill="x", pady=5)
        
        # Load clients
        clients = self.db.get_clients()
        client_names = [name for _, name, _, _ in clients]
        client_combo['values'] = client_names
        
        # Site dropdown
        ttk.Label(client_site_frame, text="Assign to Site:").pack(anchor="w", pady=(10, 0))
        site_var = tk.StringVar()
        site_combo = ttk.Combobox(client_site_frame, textvariable=site_var, width=50, state="readonly")
        site_combo.pack(fill="x", pady=5)
        
        def on_client_change(event=None):
            """Update sites when client changes"""
            client_name = client_var.get()
            if client_name:
                # Find client ID
                client_id = None
                for cid, name, _, _ in clients:
                    if name == client_name:
                        client_id = cid
                        break
                
                if client_id:
                    sites = self.db.get_sites(client_id)
                    site_names = [name for _, _, name, _, _, _ in sites]
                    site_combo['values'] = site_names
                    site_var.set("")  # Clear current selection
        
        client_combo.bind('<<ComboboxSelected>>', on_client_change)
        
        # Repository details
        details_frame = ttk.LabelFrame(dialog, text="Repository Details", padding="10")
        details_frame.pack(fill="x", padx=20, pady=10)
        
        ttk.Label(details_frame, text="Repository Name:").pack(anchor="w")
        name_var = tk.StringVar()
        ttk.Entry(details_frame, textvariable=name_var, width=50).pack(fill="x", pady=5)
        
        ttk.Label(details_frame, text="Role:").pack(anchor="w", pady=(10, 0))
        role_var = tk.StringVar(value="imported")
        role_combo = ttk.Combobox(details_frame, textvariable=role_var, width=50)
        role_combo['values'] = ["imported", "system-backup", "application-backup", "user-data", "custom"]
        role_combo.pack(fill="x", pady=5)
        
        ttk.Label(details_frame, text="Repository Password:").pack(anchor="w", pady=(10, 0))
        password_var = tk.StringVar()
        password_entry = ttk.Entry(details_frame, textvariable=password_var, show="*", width=50)
        password_entry.pack(fill="x", pady=5)
        
        # Progress frame
        progress_frame = ttk.Frame(dialog)
        progress_frame.pack(fill="x", padx=20, pady=10)
        
        progress_var = tk.StringVar(value="Ready to import...")
        ttk.Label(progress_frame, textvariable=progress_var).pack(anchor="w")
        
        progress_bar = ttk.Progressbar(progress_frame, mode='indeterminate')
        progress_bar.pack(fill="x", pady=5)
        
        # Buttons
        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill="x", padx=20, pady=20)
        
        def start_import():
            """Start the import process"""
            source_path = source_var.get().strip()
            client_name = client_var.get().strip()
            site_name = site_var.get().strip()
            repo_name = name_var.get().strip()
            role = role_var.get().strip()
            password = password_var.get().strip()
            
            # Validation
            if not source_path or not Path(source_path).exists():
                messagebox.showerror("Error", "Please select a valid source repository path")
                return
            
            if not client_name:
                messagebox.showerror("Error", "Please select a client")
                return
            
            if not site_name:
                messagebox.showerror("Error", "Please select a site")
                return
            
            if not repo_name:
                messagebox.showerror("Error", "Please enter a repository name")
                return
            
            if not password:
                messagebox.showerror("Error", "Please enter the repository password")
                return
            
            progress_bar.start()
            progress_var.set("Starting import...")
            
            def import_thread():
                try:
                    result = self.import_repository_standalone(
                        source_path, client_name, site_name, repo_name, role, password
                    )
                    
                    dialog.after(0, lambda: import_complete(result))
                    
                except Exception as e:
                    dialog.after(0, lambda: import_failed(str(e)))
            
            def import_complete(success):
                progress_bar.stop()
                if success:
                    progress_var.set("Import completed successfully!")
                    messagebox.showinfo("Import Complete", "Repository imported successfully!")
                    dialog.destroy()
                    self.refresh_images_list()  # Refresh the repository list
                else:
                    progress_var.set("Import failed!")
            
            def import_failed(error):
                progress_bar.stop()
                progress_var.set("Import failed!")
                messagebox.showerror("Import Failed", f"Failed to import repository:\n{error}")
            
            threading.Thread(target=import_thread, daemon=True).start()
        
        ttk.Button(button_frame, text="🚀 Start Import", command=start_import).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side="left", padx=5)

    def import_repository_standalone(self, source_path, client_name, site_name, repo_name, role, password):
        """Standalone repository import function"""
        try:
            self.log_step2(f"Starting import of repository: {source_path}")
            
            # Get client and site IDs
            clients = self.db.get_clients()
            client_id = None
            for cid, name, _, _ in clients:
                if name == client_name:
                    client_id = cid
                    break
            
            if not client_id:
                raise Exception(f"Client '{client_name}' not found")
            
            sites = self.db.get_sites(client_id)
            site_id = None
            for sid, _, name, _, _, _ in sites:
                if name == site_name:
                    site_id = sid
                    break
            
            if not site_id:
                raise Exception(f"Site '{site_name}' not found")
            
            # Create destination path using organized structure
            restic_base = self.get_restic_base_path()
            client_repo_dir = restic_base / client_id
            dest_repo = client_repo_dir / repo_name
            
            self.log_step2(f"Destination: {dest_repo}")
            
            # Check if destination already exists
            if dest_repo.exists():
                self.log_step2("WARNING: Destination repository already exists")
                raise Exception(f"Repository '{repo_name}' already exists for this client")
            
            # Copy repository
            self.log_step2("Copying repository files...")
            dest_repo.parent.mkdir(parents=True, exist_ok=True)
            shutil.copytree(source_path, dest_repo)
            self.log_step2("Repository files copied successfully")
            
            # Verify repository is accessible
            self.log_step2("Verifying repository integrity...")
            restic_exe = self.download_restic()
            if not restic_exe:
                raise Exception("Could not obtain restic binary")
            
            # Test repository access
            os.environ['RESTIC_REPOSITORY'] = str(dest_repo)
            os.environ['RESTIC_PASSWORD'] = password
            
            check_cmd = [restic_exe, "snapshots", "--json"]
            result = subprocess.run(check_cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
            
            if result.returncode != 0:
                # Clean up copied files on failure
                shutil.rmtree(dest_repo)
                raise Exception(f"Repository verification failed: {result.stderr}")
            
            # Parse snapshots to get count
            try:
                snapshots = json.loads(result.stdout)
                snapshot_count = len(snapshots) if snapshots else 0
                latest_snapshot = snapshots[-1]['short_id'] if snapshots else None
            except (json.JSONDecodeError, KeyError, IndexError):
                snapshot_count = 0
                latest_snapshot = None
            
            # Calculate repository size
            repo_size_gb = self.calculate_repo_size(dest_repo)
            
            # Create database entry
            image_id = generate_uuidv7()
            self.db.create_image(
                image_id=image_id,
                client_id=client_id,
                site_id=site_id,
                role=role,
                repository_path=str(dest_repo),
                repository_size_gb=repo_size_gb,
                snapshot_count=snapshot_count,
                latest_snapshot_id=latest_snapshot,
                restic_password=password
            )
            
            self.log_step2(f"Repository imported successfully with ID: {image_id}")
            self.log_step2(f"Snapshots found: {snapshot_count}")
            self.log_step2(f"Repository size: {repo_size_gb:.1f} GB")
            
            return True
            
        except Exception as e:
            self.log_step2(f"Import failed: {str(e)}")
            return False

    def scan_and_import_repository(self):
        """Scan an existing repository and create/update image records based on UUIDs"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Scan & Import Repository")
        dialog.geometry("800x700")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # Title
        ttk.Label(dialog, text="Scan Repository for Images", 
                 font=("TkDefaultFont", 14, "bold")).pack(pady=20)
        ttk.Label(dialog, text="Scan a restic repository for image UUIDs and create/update database records", 
                 font=("TkDefaultFont", 10)).pack(pady=5)
        
        # Repository selection
        repo_frame = ttk.LabelFrame(dialog, text="Repository Selection", padding="10")
        repo_frame.pack(fill="x", padx=20, pady=10)
        
        ttk.Label(repo_frame, text="Select restic repository to scan:").pack(anchor="w")
        
        repo_var = tk.StringVar()
        repo_entry = ttk.Entry(repo_frame, textvariable=repo_var, width=70)
        repo_entry.pack(fill="x", pady=5)
        
        def browse_repository():
            path = filedialog.askdirectory(title="Select Restic Repository to Scan")
            if path:
                repo_var.set(path)
        
        ttk.Button(repo_frame, text="Browse...", command=browse_repository).pack(pady=5)
        
        # Client/Site assignment for new records
        assignment_frame = ttk.LabelFrame(dialog, text="Default Assignment (for new records)", padding="10")
        assignment_frame.pack(fill="x", padx=20, pady=10)
        
        ttk.Label(assignment_frame, text="Default Client:").pack(anchor="w")
        client_var = tk.StringVar()
        client_combo = ttk.Combobox(assignment_frame, textvariable=client_var, width=50, state="readonly")
        client_combo.pack(fill="x", pady=5)
        
        # Load clients
        clients = self.db.get_clients()
        client_names = [name for _, name, _, _ in clients]
        client_combo['values'] = client_names
        
        ttk.Label(assignment_frame, text="Default Site:").pack(anchor="w", pady=(10, 0))
        site_var = tk.StringVar()
        site_combo = ttk.Combobox(assignment_frame, textvariable=site_var, width=50, state="readonly")
        site_combo.pack(fill="x", pady=5)
        
        def on_client_change(event=None):
            """Update sites when client changes"""
            client_name = client_var.get()
            if client_name:
                client_id = None
                for cid, name, _, _ in clients:
                    if name == client_name:
                        client_id = cid
                        break
                
                if client_id:
                    sites = self.db.get_sites(client_id)
                    site_names = [name for _, _, name, _, _, _ in sites]
                    site_combo['values'] = site_names
                    site_var.set("")
        
        client_combo.bind('<<ComboboxSelected>>', on_client_change)
        
        # Password
        ttk.Label(assignment_frame, text="Repository Password:").pack(anchor="w", pady=(10, 0))
        password_var = tk.StringVar()
        password_entry = ttk.Entry(assignment_frame, textvariable=password_var, show="*", width=50)
        password_entry.pack(fill="x", pady=5)
        
        # Results area
        results_frame = ttk.LabelFrame(dialog, text="Scan Results", padding="10")
        results_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        results_text = scrolledtext.ScrolledText(results_frame, height=15, font=("Consolas", 9))
        results_text.pack(fill="both", expand=True)
        
        # Progress
        progress_frame = ttk.Frame(dialog)
        progress_frame.pack(fill="x", padx=20, pady=10)
        
        progress_var = tk.StringVar(value="Ready to scan...")
        ttk.Label(progress_frame, textvariable=progress_var).pack(anchor="w")
        
        progress_bar = ttk.Progressbar(progress_frame, mode='indeterminate')
        progress_bar.pack(fill="x", pady=5)
        
        # Buttons
        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill="x", padx=20, pady=20)
        
        def start_scan():
            """Start repository scan"""
            repo_path = repo_var.get().strip()
            client_name = client_var.get().strip()
            site_name = site_var.get().strip()
            password = password_var.get().strip()
            
            # Validation
            if not repo_path or not Path(repo_path).exists():
                messagebox.showerror("Error", "Please select a valid repository path")
                return
            
            if not client_name:
                messagebox.showerror("Error", "Please select a default client")
                return
            
            if not site_name:
                messagebox.showerror("Error", "Please select a default site")
                return
            
            if not password:
                messagebox.showerror("Error", "Please enter the repository password")
                return
            
            progress_bar.start()
            progress_var.set("Scanning repository...")
            results_text.delete(1.0, tk.END)
            
            def scan_thread():
                try:
                    results = self.scan_repository_for_images(
                        repo_path, client_name, site_name, password, results_text
                    )
                    
                    dialog.after(0, lambda: scan_complete(results))
                    
                except Exception as e:
                    dialog.after(0, lambda: scan_failed(str(e)))
            
            def scan_complete(results):
                progress_bar.stop()
                progress_var.set(f"Scan completed! Found {len(results)} image records.")
                self.log_step2(f"Repository scan completed: {len(results)} image records processed")
                self.refresh_images_list()  # Refresh repository list
            
            def scan_failed(error):
                progress_bar.stop()
                progress_var.set("Scan failed!")
                results_text.insert(tk.END, f"\nERROR: {error}\n")
                self.log_step2(f"Repository scan failed: {error}")
            
            threading.Thread(target=scan_thread, daemon=True).start()
        
        ttk.Button(button_frame, text="🔍 Start Scan", command=start_scan).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Close", command=dialog.destroy).pack(side="left", padx=5)

    def scan_repository_for_images(self, repo_path, default_client_name, default_site_name, password, results_widget):
        """Scan repository for image UUIDs and create/update database records"""
        try:
            # Get default client and site IDs
            clients = self.db.get_clients()
            default_client_id = None
            for cid, name, _, _ in clients:
                if name == default_client_name:
                    default_client_id = cid
                    break
            
            if not default_client_id:
                raise Exception(f"Default client '{default_client_name}' not found")
            
            sites = self.db.get_sites(default_client_id)
            default_site_id = None
            for sid, _, name, _, _, _ in sites:
                if name == default_site_name:
                    default_site_id = sid
                    break
            
            if not default_site_id:
                raise Exception(f"Default site '{default_site_name}' not found")
            
            # Setup restic environment
            restic_exe = self.download_restic()
            if not restic_exe:
                raise Exception("Could not obtain restic binary")
            
            os.environ['RESTIC_REPOSITORY'] = str(repo_path)
            os.environ['RESTIC_PASSWORD'] = password
            
            results_widget.insert(tk.END, f"Scanning repository: {repo_path}\n")
            results_widget.insert(tk.END, f"Default assignment: {default_client_name} / {default_site_name}\n\n")
            results_widget.update()
            
            # Get all snapshots with tags
            self.log_step2("Querying repository snapshots...")
            snapshots_cmd = [restic_exe, "snapshots", "--json"]
            result = subprocess.run(snapshots_cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
            
            if result.returncode != 0:
                raise Exception(f"Failed to query snapshots: {result.stderr}")
            
            snapshots = json.loads(result.stdout) if result.stdout.strip() else []
            results_widget.insert(tk.END, f"Found {len(snapshots)} snapshots in repository\n\n")
            
            # Extract image UUIDs from snapshot tags
            image_uuids = set()
            for snapshot in snapshots:
                tags = snapshot.get('tags', [])
                for tag in tags:
                    if tag.startswith('image-uuid-'):
                        uuid_str = tag.replace('image-uuid-', '')
                        image_uuids.add(uuid_str)
                        results_widget.insert(tk.END, f"Found image UUID: {uuid_str}\n")
                        results_widget.update()
            
            results_widget.insert(tk.END, f"\nTotal unique image UUIDs found: {len(image_uuids)}\n\n")
            results_widget.update()
            
            # Check each UUID against database
            processed_records = []
            for image_uuid in image_uuids:
                results_widget.insert(tk.END, f"Processing UUID: {image_uuid}\n")
                results_widget.update()
                
                # Check if record exists
                existing_images = self.db.get_images()
                existing_record = None
                for image in existing_images:
                    if image[0] == image_uuid:  # image[0] is the ID
                        existing_record = image
                        break
                
                if existing_record:
                    # Update existing record if missing role information
                    role = existing_record[4] if len(existing_record) > 4 else None
                    if not role or role.strip() == "":
                        # Update with default role
                        with sqlite3.connect(self.db.db_path) as conn:
                            cursor = conn.cursor()
                            cursor.execute('''
                                UPDATE images SET role = ?, updated_at = CURRENT_TIMESTAMP 
                                WHERE id = ?
                            ''', ("system-backup", image_uuid))
                            conn.commit()
                        
                        results_widget.insert(tk.END, f"  ✓ Updated role for existing record\n")
                        self.log_step2(f"Updated role for image {image_uuid}")
                    else:
                        results_widget.insert(tk.END, f"  ✓ Record exists with role: {role}\n")
                
                else:
                    # Create new record
                    try:
                        self.db.create_image(
                            image_id=image_uuid,
                            client_id=default_client_id,
                            site_id=default_site_id,
                            role="system-backup",
                            repository_path=str(repo_path),
                            repository_size_gb=0,  # Will be calculated later
                            snapshot_count=0,      # Will be calculated later
                            latest_snapshot_id=None,
                            restic_password=password
                        )
                        
                        results_widget.insert(tk.END, f"  ✓ Created new record\n")
                        self.log_step2(f"Created new image record for UUID {image_uuid}")
                        
                    except Exception as e:
                        results_widget.insert(tk.END, f"  ✗ Failed to create record: {e}\n")
                        self.log_step2(f"Failed to create record for UUID {image_uuid}: {e}")
                
                processed_records.append(image_uuid)
                results_widget.update()
            
            # Update repository statistics
            repo_size_gb = self.calculate_repo_size(Path(repo_path))
            total_snapshots = len(snapshots)
            
            results_widget.insert(tk.END, f"\nRepository Statistics:\n")
            results_widget.insert(tk.END, f"  Size: {repo_size_gb:.1f} GB\n")
            results_widget.insert(tk.END, f"  Total Snapshots: {total_snapshots}\n")
            results_widget.insert(tk.END, f"  Image Records: {len(processed_records)}\n\n")
            
            # Update all records for this repository with current stats
            with sqlite3.connect(self.db.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute('''
                    UPDATE images 
                    SET repository_size_gb = ?, snapshot_count = ?, updated_at = CURRENT_TIMESTAMP
                    WHERE repository_path = ?
                ''', (repo_size_gb, total_snapshots, str(repo_path)))
                conn.commit()
            
            results_widget.insert(tk.END, "✓ Repository scan completed successfully!\n")
            self.log_step2(f"Repository scan completed: {len(processed_records)} records processed")
            
            return processed_records
            
        except Exception as e:
            results_widget.insert(tk.END, f"\nERROR: {str(e)}\n")
            raise e

    def export_database(self):
        """Export database to portable format"""
        try:
            export_path = filedialog.asksaveasfilename(
                title="Export Database",
                defaultextension=".json",
                filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")]
            )
            
            if not export_path:
                return
            
            # Export all data
            export_data = {
                "version": "1.0",
                "exported_at": datetime.now().isoformat(),
                "clients": self.db.get_clients(),
                "sites": self.db.get_sites(),
                "images": self.db.get_images(),
                "config": {}
            }
            
            # Get config data
            with sqlite3.connect(self.db.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT key, value FROM config")
                export_data["config"] = dict(cursor.fetchall())
            
            with open(export_path, 'w') as f:
                json.dump(export_data, f, indent=2, default=str)
            
            self.log(f"SUCCESS: Database exported to {export_path}")
            messagebox.showinfo("Export Complete", f"Database exported to:\n{export_path}")
            
        except Exception as e:
            self.log(f"ERROR: Database export failed: {e}")
            messagebox.showerror("Export Failed", f"Failed to export database:\n{e}")

    def import_database(self):
        """Import database from portable format"""
        try:
            import_path = filedialog.askopenfilename(
                title="Import Database",
                filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")]
            )
            
            if not import_path:
                return
            
            if not messagebox.askyesno("Confirm Import", 
                "This will merge the imported data with your existing database.\n"
                "Existing records with the same IDs will be updated.\n\n"
                "Continue?"):
                return
            
            with open(import_path, 'r') as f:
                import_data = json.load(f)
            
            # Validate import data
            required_keys = ["clients", "sites", "images"]
            for key in required_keys:
                if key not in import_data:
                    raise ValueError(f"Import file missing required key: {key}")
            
            # Import data
            imported_counts = {"clients": 0, "sites": 0, "images": 0}
            
            with sqlite3.connect(self.db.db_path) as conn:
                cursor = conn.cursor()
                
                # Import clients
                for client_data in import_data["clients"]:
                    try:
                        cursor.execute('''
                            INSERT OR REPLACE INTO clients (id, name, short_name, description)
                            VALUES (?, ?, ?, ?)
                        ''', client_data[:4])
                        imported_counts["clients"] += 1
                    except Exception as e:
                        self.log(f"WARNING: Failed to import client {client_data[0]}: {e}")
                
                # Import sites
                for site_data in import_data["sites"]:
                    try:
                        cursor.execute('''
                            INSERT OR REPLACE INTO sites (id, client_id, name, short_name, description)
                            VALUES (?, ?, ?, ?, ?)
                        ''', site_data[:5])
                        imported_counts["sites"] += 1
                    except Exception as e:
                        self.log(f"WARNING: Failed to import site {site_data[0]}: {e}")
                
                # Import images
                for image_data in import_data["images"]:
                    try:
                        cursor.execute('''
                            INSERT OR REPLACE INTO images 
                            (id, client_id, site_id, role, wim_source_path, vhdx_path, vhdx_size_gb, vm_name, vm_created, status)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ''', image_data[:10])
                        imported_counts["images"] += 1
                    except Exception as e:
                        self.log(f"WARNING: Failed to import image {image_data[0]}: {e}")
                
                # Import config
                if "config" in import_data:
                    for key, value in import_data["config"].items():
                        try:
                            cursor.execute('''
                                INSERT OR REPLACE INTO config (key, value)
                                VALUES (?, ?)
                            ''', (key, value))
                        except Exception as e:
                            self.log(f"WARNING: Failed to import config {key}: {e}")
                
                conn.commit()
            
            self.log(f"SUCCESS: Database import completed")
            self.log(f"  Clients: {imported_counts['clients']}")
            self.log(f"  Sites: {imported_counts['sites']}")
            self.log(f"  Images: {imported_counts['images']}")
            
            # Refresh UI
            self.refresh_client_site_data()
            self.refresh_images_list()
            
            messagebox.showinfo("Import Complete", 
                f"Database import completed:\n"
                f"• Clients: {imported_counts['clients']}\n"
                f"• Sites: {imported_counts['sites']}\n"
                f"• Images: {imported_counts['images']}")
            
        except Exception as e:
            self.log(f"ERROR: Database import failed: {e}")
            messagebox.showerror("Import Failed", f"Failed to import database:\n{e}")

    def backup_database(self):
        """Create backup of database"""
        try:
            backup_path = filedialog.asksaveasfilename(
                title="Backup Database",
                defaultextension=".db",
                filetypes=[("Database Files", "*.db"), ("All Files", "*.*")],
                initialfile=f"pyc_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
            )
            
            if not backup_path:
                return
            
            shutil.copy2(self.db.db_path, backup_path)
            self.log(f"SUCCESS: Database backed up to {backup_path}")
            messagebox.showinfo("Backup Complete", f"Database backed up to:\n{backup_path}")
            
        except Exception as e:
            self.log(f"ERROR: Database backup failed: {e}")
            messagebox.showerror("Backup Failed", f"Failed to backup database:\n{e}")

    def clean_orphaned_records(self):
        """Clean orphaned database records"""
        try:
            if not messagebox.askyesno("Confirm Cleanup", 
                "This will remove database records for images that no longer exist on disk.\n"
                "This cannot be undone.\n\n"
                "Continue?"):
                return
            
            cleaned_count = 0
            images = self.db.get_images()
            
            with sqlite3.connect(self.db.db_path) as conn:
                cursor = conn.cursor()
                
                for image_data in images:
                    image_id = image_data[0]
                    vhdx_path = image_data[2]
                    
                    if vhdx_path and not Path(vhdx_path).exists():
                        cursor.execute("DELETE FROM images WHERE id = ?", (image_id,))
                        cleaned_count += 1
                        self.log(f"INFO: Removed orphaned record for {image_id}")
                
                conn.commit()
            
            self.log(f"SUCCESS: Cleaned {cleaned_count} orphaned records")
            self.refresh_images_list()
            messagebox.showinfo("Cleanup Complete", f"Cleaned {cleaned_count} orphaned records")
            
        except Exception as e:
            self.log(f"ERROR: Cleanup failed: {e}")
            messagebox.showerror("Cleanup Failed", f"Failed to clean orphaned records:\n{e}")

    def show_database_stats(self):
        """Show database statistics"""
        try:
            with sqlite3.connect(self.db.db_path) as conn:
                cursor = conn.cursor()
                
                # Get counts
                cursor.execute("SELECT COUNT(*) FROM clients")
                client_count = cursor.fetchone()[0]
                
                cursor.execute("SELECT COUNT(*) FROM sites")
                site_count = cursor.fetchone()[0]
                
                cursor.execute("SELECT COUNT(*) FROM images")
                image_count = cursor.fetchone()[0]
                
                cursor.execute("SELECT COUNT(*) FROM images WHERE vm_created = 1")
                vm_count = cursor.fetchone()[0]
                
                cursor.execute("SELECT SUM(vhdx_size_gb) FROM images")
                total_size = cursor.fetchone()[0] or 0
                
                # Get database size
                db_size = self.db.db_path.stat().st_size / (1024**2)  # MB
                
                stats_text = f"""Database Statistics:

• Clients: {client_count}
• Sites: {site_count}
• Images: {image_count}
• VMs Created: {vm_count}
• Total VHDX Size: {total_size} GB
• Database Size: {db_size:.1f} MB
• Database Location: {self.db.db_path}
• Image Storage: {self.image_store_path}"""
                
                messagebox.showinfo("Database Statistics", stats_text)
                
        except Exception as e:
            self.log(f"ERROR: Failed to get database stats: {e}")
            messagebox.showerror("Error", f"Failed to get database statistics:\n{e}")

    def on_step2_tab_changed(self, event=None):
        """Handle Step 2 tab changes to update data"""
        try:
            selected_tab = self.step2_notebook.select()
            tab_text = self.step2_notebook.tab(selected_tab, "text")
            
            if "📊 Dashboard" in tab_text:
                self.update_dashboard_stats()
            elif "📁 Browse" in tab_text:
                self.refresh_images_list()
            elif "🗄️ Database" in tab_text:
                self.refresh_client_site_data()
        except Exception as e:
            # Fail silently for tab change events
            pass

    def update_dashboard_stats(self):
        """Update dashboard statistics"""
        try:
            with sqlite3.connect(self.db.db_path) as conn:
                cursor = conn.cursor()
                
                # Get counts
                cursor.execute("SELECT COUNT(*) FROM clients")
                client_count = cursor.fetchone()[0]
                
                cursor.execute("SELECT COUNT(*) FROM sites")
                site_count = cursor.fetchone()[0]
                
                cursor.execute("SELECT COUNT(*) FROM images")
                image_count = cursor.fetchone()[0]
                
                cursor.execute("SELECT COUNT(*) FROM images WHERE vm_created = 1")
                vm_count = cursor.fetchone()[0]
                
                cursor.execute("SELECT SUM(vhdx_size_gb) FROM images")
                total_size = cursor.fetchone()[0] or 0
                
                # Calculate storage used
                storage_used_gb = 0
                try:
                    for image_file in self.image_store_path.glob("*.wim"):
                        storage_used_gb += image_file.stat().st_size / (1024**3)
                    for image_file in self.image_store_path.glob("*.vhdx"):
                        storage_used_gb += image_file.stat().st_size / (1024**3)
                except:
                    storage_used_gb = total_size  # Fallback to database values
                
                # Update labels
                if hasattr(self, 'stats_labels'):
                    self.stats_labels["Total Images"].config(text=str(image_count))
                    self.stats_labels["Total VMs"].config(text=str(vm_count))
                    self.stats_labels["Total Clients"].config(text=str(client_count))
                    self.stats_labels["Total Sites"].config(text=str(site_count))
                    self.stats_labels["Storage Used"].config(text=f"{storage_used_gb:.1f} GB")
                
        except Exception as e:
            self.log(f"ERROR: Failed to update dashboard stats: {e}")

    def populate_step3_frame(self):
        """Step 3: Generalize and Cleanup"""
        frame = self.step_frames[3]
        
        # Step description
        desc_frame = ttk.LabelFrame(frame, text="Step 3: Generalize and Cleanup", padding="10")
        desc_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(desc_frame, text="Prepare the system for cloning using Sysprep.", 
                 font=("TkDefaultFont", 9)).pack()

        # Configuration
        config_frame = ttk.LabelFrame(frame, text="Configuration", padding="10")
        config_frame.pack(fill="x", pady=(0, 10))
        config_frame.columnconfigure(0, weight=1)

        # Audit Mode Status Indicator
        audit_status_frame = ttk.Frame(config_frame)
        audit_status_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        audit_status_frame.columnconfigure(1, weight=1)
        
        ttk.Label(audit_status_frame, text="Audit Mode Status:", font=("TkDefaultFont", 9, "bold")).grid(row=0, column=0, sticky="w")
        self.audit_status_label = ttk.Label(audit_status_frame, text="Checking...", foreground="orange")
        self.audit_status_label.grid(row=0, column=1, sticky="w", padx=(10, 0))
        
        # Check audit mode status on startup
        self.update_audit_mode_status()

        self.skip_user_cleanup_var = tk.BooleanVar()
        self.skip_agent_cleanup_var = tk.BooleanVar()
        self.skip_log_cleanup_var = tk.BooleanVar()

        ttk.Checkbutton(config_frame, text="Skip User Profile & Account Cleanup", variable=self.skip_user_cleanup_var).grid(row=1, column=0, sticky="w")
        ttk.Checkbutton(config_frame, text="Skip Agent Identity Cleanup (e.g., NinjaRMM)", variable=self.skip_agent_cleanup_var).grid(row=2, column=0, sticky="w")
        ttk.Checkbutton(config_frame, text="Skip System Log Cleanup (Event Logs, etc.)", variable=self.skip_log_cleanup_var).grid(row=3, column=0, sticky="w")
        
        self.generalize_button = ttk.Button(config_frame, text="Prepare and Generalize System", command=self.start_generalization_thread)
        self.generalize_button.grid(row=4, column=0, sticky="ew", pady=10)

    def populate_step4_frame(self):
        """Step 4: Capture into WIM"""
        frame = self.step_frames[4]
        
        # Step description
        desc_frame = ttk.LabelFrame(frame, text="Step 4: Capture into WIM", padding="10")
        desc_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(desc_frame, text="Capture the generalized system into a WIM file.", 
                 font=("TkDefaultFont", 9)).pack()

        # Configuration
        config_frame = ttk.LabelFrame(frame, text="Configuration", padding="10")
        config_frame.pack(fill="x", pady=(0, 10))
        config_frame.columnconfigure(1, weight=1)
        
        # VHDX Path
        ttk.Label(config_frame, text="VHDX Path:").grid(row=0, column=0, sticky="w", pady=2)
        self.vhdx_path_var = tk.StringVar()
        self.vhdx_path_entry = ttk.Entry(config_frame, textvariable=self.vhdx_path_var)
        self.vhdx_path_entry.grid(row=0, column=1, sticky="we", padx=5)
        self.vhdx_browse_button = ttk.Button(config_frame, text="Browse...", command=self.browse_vhdx_file)
        self.vhdx_browse_button.grid(row=0, column=2, sticky="e")
        
        # gptgen Path
        ttk.Label(config_frame, text="gptgen Path:").grid(row=1, column=0, sticky="w", pady=2)
        self.gptgen_path_var = tk.StringVar(value=str(Path("./gptgen.exe").resolve()))
        self.gptgen_entry = ttk.Entry(config_frame, textvariable=self.gptgen_path_var, state='readonly')
        self.gptgen_entry.grid(row=1, column=1, sticky="we", padx=5)
        self.gptgen_find_button = ttk.Button(config_frame, text="Help Me Find It", command=self.open_gptgen_download_page)
        self.gptgen_find_button.grid(row=1, column=2, sticky="e")
        
        # Action Button
        self.capture_button = ttk.Button(frame, text="Capture WIM", command=self.start_wim_capture_thread)
        self.capture_button.pack(pady=10, fill="x")

    def populate_step5_frame(self):
        """Step 5: Deploy WIM"""
        frame = self.step_frames[5]
        
        # Step description
        desc_frame = ttk.LabelFrame(frame, text="Step 5: Deploy WIM", padding="10")
        desc_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(desc_frame, text="Deploy the captured WIM file to target machines.", 
                 font=("TkDefaultFont", 9)).pack()

        # Configuration
        config_frame = ttk.LabelFrame(frame, text="Configuration", padding="10")
        config_frame.pack(fill="x", pady=(0, 10))
        config_frame.columnconfigure(1, weight=1)

        # WIM Path
        ttk.Label(config_frame, text="WIM Path:").grid(row=0, column=0, sticky="w", pady=2)
        self.wim_path_var = tk.StringVar()
        self.wim_path_entry = ttk.Entry(config_frame, textvariable=self.wim_path_var)
        self.wim_path_entry.grid(row=0, column=1, sticky="we", padx=5)
        self.browse_wim_button = ttk.Button(config_frame, text="Browse...", command=self.browse_wim_file)
        self.browse_wim_button.grid(row=0, column=2, sticky="e")
        
        # Action Button
        self.deploy_button = ttk.Button(frame, text="Deploy WIM", command=self.start_wim_deployment_thread)
        self.deploy_button.pack(pady=10, fill="x")

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



    def browse_vhdx_file(self):
        """Opens a file dialog to select a VHDX file."""
        path = filedialog.askopenfilename(
            title="Select VHDX File",
            filetypes=(("VHDX Files", "*.vhdx"), ("All files", "*.*"))
        )
        if path:
            self.vhdx_path_var.set(path)


    def browse_wim_file(self):
        """Opens a file dialog to select a WIM file."""
        path = filedialog.askopenfilename(
            title="Select WIM File",
            filetypes=(("WIM Files", "*.wim"), ("All files", "*.*"))
        )
        if path:
            self.wim_path_var.set(path)

    def open_gptgen_download_page(self):
        """Opens a web browser to the gptgen download page."""
        url = "https://sourceforge.net/projects/gptgen/"
        self.log(f"INFO: Opening web browser to {url} for gptgen.")
        self.log("INFO: Please download the tool, extract it, and place 'gptgen.exe' in the same folder as this application.")
        webbrowser.open(url)

    def get_available_space(self, path):
        """Get available disk space in bytes for the given path."""
        return shutil.disk_usage(path).free

    def get_drive_size(self, drive_letter):
        """Get total size of a drive in bytes."""
        if not drive_letter.endswith(':'):
            drive_letter += ':'
        if not drive_letter.endswith('\\'):
            drive_letter += '\\'
        return shutil.disk_usage(drive_letter).total

    def install_to_public_desktop(self):
        """Copies this program to the Public Desktop as OIP.exe for easy access."""
        try:
            # Determine the current executable path
            if getattr(sys, 'frozen', False):
                # Running as compiled executable
                current_exe = sys.executable
                self.log("INFO: Detected compiled executable.")
            else:
                # Running as Python script - need to inform user
                self.log("ERROR: This feature only works with compiled executables (.exe files).")
                self.log("ERROR: Currently running as Python script. Please compile to .exe first.")
                messagebox.showerror("Compilation Required", 
                    "This feature only works with compiled executables.\n\n"
                    "Please compile this Python script to an .exe file using:\n"
                    "- PyInstaller: pyinstaller --onefile --windowed windows_image_prep_gui.py\n"
                    "- cx_Freeze or other Python-to-exe tools")
                return False

            # Define destination path
            system_drive = os.environ.get('SystemDrive', 'C:')
            public_desktop = Path(system_drive) / "Users" / "Public" / "Desktop"
            destination = public_desktop / "OIP.exe"
            
            self.log(f"INFO: Installing to: {destination}")
            
            # Check if Public Desktop exists
            if not public_desktop.exists():
                self.log(f"ERROR: Public Desktop folder not found: {public_desktop}")
                messagebox.showerror("Installation Failed", f"Public Desktop folder not found:\n{public_desktop}")
                return False
            
            # Check if file already exists
            if destination.exists():
                if not messagebox.askyesno("File Exists", 
                    f"OIP.exe already exists on the Public Desktop.\n\nOverwrite it?"):
                    return False
                try:
                    destination.unlink()
                    self.log("INFO: Removed existing OIP.exe")
                except Exception as e:
                    self.log(f"ERROR: Could not remove existing file: {e}")
                    messagebox.showerror("Installation Failed", f"Could not remove existing file:\n{e}")
                    return False
            
            # Copy the executable
            self.log("INFO: Copying executable to Public Desktop...")
            import shutil
            shutil.copy2(current_exe, destination)
            
            # Verify the copy was successful
            if destination.exists():
                file_size = destination.stat().st_size
                self.log(f"SUCCESS: OIP.exe installed to Public Desktop ({file_size:,} bytes)")
                self.log("SUCCESS: All users can now access the tool from their desktop.")
                messagebox.showinfo("Installation Complete", 
                    f"Successfully installed OIP.exe to:\n{destination}\n\n"
                                        "The tool is now available on all user desktops.")
                return True
            else:
                self.log("ERROR: File copy appeared to succeed but destination file not found.")
                return False
                
        except Exception as e:
            self.log(f"ERROR: Installation failed: {e}")
            messagebox.showerror("Installation Failed", f"Could not install to Public Desktop:\n{e}")
            return False

    def check_audit_mode(self):
        """Check if Windows is currently in audit mode."""
        try:
            # Check the Setup State registry key
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 
                                  r"SOFTWARE\Microsoft\Windows\CurrentVersion\Setup\State") as key:
                    image_state, _ = winreg.QueryValueEx(key, "ImageState")
                    self.log(f"INFO: Windows ImageState: {image_state}")
                    
                    # IMAGE_STATE_COMPLETE = "IMAGE_STATE_COMPLETE"
                    # IMAGE_STATE_UNDEPLOYABLE = "IMAGE_STATE_UNDEPLOYABLE" (audit mode)
                    if image_state == "IMAGE_STATE_UNDEPLOYABLE":
                        return True
                    else:
                        return False
            except (FileNotFoundError, OSError):
                self.log("WARNING: Could not read Setup State registry key.")
                
            # Fallback: Check for audit mode specific registry entries
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                  r"SYSTEM\Setup\Status") as key:
                    try:
                        audit_boot, _ = winreg.QueryValueEx(key, "AuditBoot")
                        return audit_boot == 1
                    except FileNotFoundError:
                        pass
            except (FileNotFoundError, OSError):
                pass
                
            # If we can't determine, assume not in audit mode for safety
            return False
            
        except Exception as e:
            self.log(f"ERROR: Failed to check audit mode: {e}")
            return False

    def show_audit_mode_warning(self):
        """Show comprehensive warning about not being in audit mode."""
        warning_text = """⚠️  CRITICAL WARNING: NOT IN AUDIT MODE  ⚠️

You are about to run Sysprep Generalize on a system that is NOT in Audit Mode.

🔥 THIS IS EXTREMELY DANGEROUS AND WILL:
   • Remove all user accounts and profiles (except built-in accounts)
   • Delete user data and personalization
   • Reset Windows activation
   • Make the system non-bootable for current users
   • Require complete reconfiguration after reboot

📋 AUDIT MODE is the SAFE way to prepare images:
   • Boot to audit mode: Ctrl+Shift+F3 during OOBE
   • Or run: sysprep /audit /reboot
   • Then run this generalization tool

🛡️  RECOMMENDED ACTIONS:
   1. STOP NOW and reboot to audit mode first
   2. Or ensure this is a disposable test system
   3. Or create a full system backup before continuing

⚠️  DO NOT CONTINUE ON PRODUCTION SYSTEMS  ⚠️

Are you absolutely certain you want to proceed?
This action cannot be undone!"""

        # Create a custom dialog with more prominent warning
        import tkinter as tk
        from tkinter import messagebox
        
        result = messagebox.askyesno(
            "🔥 CRITICAL WARNING - NOT IN AUDIT MODE 🔥",
            warning_text,
            icon='warning',
            default='no'
        )
        
        if result:
            # Double confirmation for extra safety
            final_warning = """FINAL CONFIRMATION REQUIRED

You have chosen to proceed despite the audit mode warning.

This will PERMANENTLY ALTER your Windows installation.

Type 'I UNDERSTAND THE RISK' below to confirm:"""
            
            # Create a simple input dialog
            confirmation_window = tk.Toplevel(self.root)
            confirmation_window.title("Final Confirmation Required")
            confirmation_window.geometry("400x200")
            confirmation_window.transient(self.root)
            confirmation_window.grab_set()
            
            # Center the window
            confirmation_window.update_idletasks()
            x = (confirmation_window.winfo_screenwidth() // 2) - (confirmation_window.winfo_width() // 2)
            y = (confirmation_window.winfo_screenheight() // 2) - (confirmation_window.winfo_height() // 2)
            confirmation_window.geometry(f"+{x}+{y}")
            
            ttk.Label(confirmation_window, text=final_warning, wraplength=380).pack(pady=10, padx=10)
            
            entry_var = tk.StringVar()
            entry = ttk.Entry(confirmation_window, textvariable=entry_var, width=30)
            entry.pack(pady=10)
            entry.focus()
            
            confirmed = [False]  # Use list to allow modification in nested function
            
            def check_confirmation():
                if entry_var.get().strip().upper() == "I UNDERSTAND THE RISK":
                    confirmed[0] = True
                    confirmation_window.destroy()
                else:
                    messagebox.showerror("Incorrect Confirmation", "You must type exactly: I UNDERSTAND THE RISK")
            
            def cancel_confirmation():
                confirmation_window.destroy()
            
            button_frame = ttk.Frame(confirmation_window)
            button_frame.pack(pady=10)
            
            ttk.Button(button_frame, text="Confirm", command=check_confirmation).pack(side="left", padx=5)
            ttk.Button(button_frame, text="Cancel (Recommended)", command=cancel_confirmation).pack(side="left", padx=5)
            
            # Bind Enter key to confirmation
            entry.bind('<Return>', lambda e: check_confirmation())
            
            confirmation_window.wait_window()
            return confirmed[0]
        
        return False

    def update_audit_mode_status(self):
        """Updates the audit mode status label."""
        is_audit_mode = self.check_audit_mode()
        if is_audit_mode:
            self.audit_status_label.config(text="✅ Audit Mode Active (Safe)", foreground="green")
        else:
            self.audit_status_label.config(text="⚠️ Audit Mode Inactive (Risky)", foreground="red")

    def start_image_creation_thread(self):
        """Legacy method - disk2vhd functionality removed. Use VSS+DISM method instead."""
        self.log("ERROR: This method is deprecated. Please use the VSS+DISM method in Step 1.")
        messagebox.showerror("Deprecated Method", "This method has been removed. Please use the 'Create WIM Image (VSS + DISM)' button instead.")

    def create_image_worker(self):
        """Legacy disk2vhd worker - functionality removed. Use VSS+DISM method instead."""
        self.log("ERROR: Legacy disk2vhd worker called - this functionality has been removed.")
        self.log("INFO: Please use the VSS+DISM method in Step 1 instead.")

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
        
        # First, check if we're in audit mode
        self.log("INFO: Checking if Windows is in Audit Mode...")
        is_audit_mode = self.check_audit_mode()
        
        if is_audit_mode:
            self.log("SUCCESS: System is in Audit Mode - safe to proceed with generalization.")
        else:
            self.log("WARNING: System is NOT in Audit Mode!")
            self.log("WARNING: Running Sysprep Generalize outside of Audit Mode is dangerous!")
            
            # Show the comprehensive warning and get user confirmation
            if not self.show_audit_mode_warning():
                self.log("INFO: User cancelled generalization due to audit mode warning.")
                return
            
            self.log("WARNING: User confirmed to proceed despite not being in Audit Mode.")
        
        # Original confirmation dialog
        if not messagebox.askyesno("Confirm Generalization", 
                                   "This will prepare the system for cloning using Sysprep. The process is irreversible and will SHUT DOWN the computer upon completion. Are you sure you want to continue?"):
            return
        
        self.generalize_button.config(state="disabled")
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

    # Note: Old VHDX processing methods removed - functionality replaced by new Step 2 professional image creation

    def start_wim_capture_thread(self):
        """Starts the WIM capture process in a new thread."""
        # Basic validation
        if not hasattr(self, 'vhdx_path_var') or not self.vhdx_path_var.get():
            messagebox.showerror("Missing VHDX", "Please specify a VHDX file to capture from.")
            return

        if not messagebox.askyesno("Confirm WIM Capture", 
                                   "This will capture the VHDX into a WIM file. Continue?"):
            return
            
        self.capture_button.config(state="disabled")
        thread = threading.Thread(target=self.wim_capture_worker)
        thread.daemon = True
        thread.start()

    def wim_capture_worker(self):
        """Worker function for WIM capture."""
        try:
            self.log("--- Starting WIM Capture ---")
            self.log("INFO: This is a placeholder for WIM capture functionality.")
            self.log("INFO: Implementation would use DISM to capture the VHDX into WIM format.")
            self.log("INFO: Feature coming soon!")
            
        except Exception as e:
            self.log(f"FATAL: WIM capture failed: {e}")
        finally:
            self.capture_button.config(state="normal")

    def start_wim_deployment_thread(self):
        """Starts the WIM deployment process in a new thread."""
        # Basic validation
        if not hasattr(self, 'wim_path_var') or not self.wim_path_var.get():
            messagebox.showerror("Missing WIM", "Please specify a WIM file to deploy.")
            return
            
        if not messagebox.askyesno("Confirm WIM Deployment", 
                                   "This will deploy the WIM file to the target system. Continue?"):
            return
        
        self.deploy_button.config(state="disabled")
        thread = threading.Thread(target=self.wim_deployment_worker)
        thread.daemon = True
        thread.start()

    def wim_deployment_worker(self):
        """Worker function for WIM deployment."""
        try:
            self.log("--- Starting WIM Deployment ---")
            self.log("INFO: This is a placeholder for WIM deployment functionality.")
            self.log("INFO: Implementation would use DISM to apply the WIM to target disks.")
            self.log("INFO: Feature coming soon!")
            
        except Exception as e:
            self.log(f"FATAL: WIM deployment failed: {e}")
        finally:
            self.deploy_button.config(state="normal")

    def check_vss_prerequisites(self):
        """Check system prerequisites for successful VSS capture."""
        self.log("INFO: Checking VSS prerequisites...")
        
        # Check available memory
        try:
            import psutil
            memory = psutil.virtual_memory()
            available_gb = memory.available / (1024**3)
            self.log(f"INFO: Available RAM: {available_gb:.1f} GB")
            if available_gb < 2:
                self.log("WARNING: Low available RAM may cause VSS issues")
                self.log("RECOMMENDATION: Close applications to free up memory")
        except ImportError:
            self.log("INFO: psutil not available - cannot check memory")
        
        # Check for running backup services that might interfere
        interfering_services = ["BackupExecVSSProvider", "VeeamVSSSupport", "VSS"]
        for service in interfering_services:
            try:
                result = subprocess.run(["sc", "query", service], 
                                      capture_output=True, text=True)
                if "RUNNING" in result.stdout:
                    self.log(f"INFO: Service {service} is running")
                    if service != "VSS":
                        self.log(f"WARNING: {service} may interfere with VSS")
            except:
                pass
        
        # Check disk space for local repository storage
        repo_type = self.repo_type_var.get() if hasattr(self, 'repo_type_var') else "local"
        if repo_type == "local":
            try:
                restic_base = self.get_restic_base_path()
                if restic_base.exists():
                    free_space = shutil.disk_usage(restic_base).free / (1024**3)
                    self.log(f"INFO: Free space at repository location: {free_space:.1f} GB")
                    if free_space < 50:
                        self.log("WARNING: Low disk space may cause backup to fail")
                        self.log("RECOMMENDATION: Free up disk space or configure S3 storage")
            except Exception as e:
                self.log(f"WARNING: Could not check disk space: {e}")
        else:
            self.log("INFO: Using S3 storage - local disk space check skipped")
        
        self.log("INFO: VSS prerequisites check completed")

    def create_vss_wim_image(self):
        """Creates a WIM image using VSS shadow copy + DISM (safe approach)."""
        # This method is deprecated - use the modern restic backup workflow instead
        self.log("ERROR: This DISM method is deprecated. Use the modern Restic backup in Step 1.")
        messagebox.showerror("Method Deprecated", 
                           "DISM-based WIM creation has been replaced with the modern Restic workflow in Step 1.\n\n" +
                           "Please use 'Create System Backup' in Step 1 instead.")
        return False

    def create_vss_shadow_copy(self, drive_letter):
        """Create a VSS shadow copy of the specified drive."""
        try:
            self.log(f"INFO: Creating VSS shadow copy of drive {drive_letter}")
            
            # First check if VSS service is running
            self.log("INFO: Checking Volume Shadow Copy service status...")
            try:
                vss_status_proc = subprocess.run(
                    ["sc", "query", "VSS"], 
                    capture_output=True, text=True, encoding='utf-8', errors='ignore'
                )
                if "RUNNING" not in vss_status_proc.stdout:
                    self.log("WARNING: Volume Shadow Copy service is not running")
                    self.log("INFO: Attempting to start VSS service...")
                    start_proc = subprocess.run(
                        ["sc", "start", "VSS"], 
                        capture_output=True, text=True, encoding='utf-8', errors='ignore'
                    )
                    if start_proc.returncode != 0:
                        self.log(f"ERROR: Failed to start VSS service: {start_proc.stderr}")
                        
                        # Try alternative VSS service startup
                        self.log("INFO: Trying alternative VSS service startup...")
                        start_proc2 = subprocess.run(
                            ["net", "start", "VSS"], 
                            capture_output=True, text=True, encoding='utf-8', errors='ignore'
                        )
                        if start_proc2.returncode != 0:
                            self.log(f"ERROR: Alternative VSS startup also failed: {start_proc2.stderr}")
                else:
                    self.log("SUCCESS: VSS service is running")
            except Exception as e:
                self.log(f"WARNING: Could not check VSS service status: {e}")
            
            # Use modern PowerShell approach first (more reliable)
            self.log("INFO: Attempting VSS creation using PowerShell (modern approach)...")
            try:
                ps_cmd = f"""
                $volume = Get-WmiObject -List Win32_ShadowCopy
                $result = $volume.Create('{drive_letter}\\', 'ClientAccessible')
                if ($result.ReturnValue -eq 0) {{
                    $shadowId = $result.ShadowID
                    Write-Host "POWERSHELL_SUCCESS:$shadowId"
                }} else {{
                    Write-Host "POWERSHELL_ERROR:$($result.ReturnValue)"
                }}
                """
                
                ps_proc = subprocess.run([
                    "powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_cmd
                ], capture_output=True, text=True, encoding='utf-8', errors='ignore')
                
                if ps_proc.returncode == 0 and "POWERSHELL_SUCCESS:" in ps_proc.stdout:
                    shadow_id = ps_proc.stdout.split("POWERSHELL_SUCCESS:")[1].strip()
                    self.log(f"SUCCESS: PowerShell created shadow copy with ID: {shadow_id}")
                    return shadow_id
                else:
                    self.log(f"WARNING: PowerShell VSS creation failed: {ps_proc.stdout} {ps_proc.stderr}")
            except Exception as e:
                self.log(f"WARNING: PowerShell VSS approach failed: {e}")
            
            # Fallback to vssadmin with fixed syntax
            self.log("INFO: Falling back to vssadmin command...")
            vss_cmd = ["vssadmin", "create", "shadow", f"/for={drive_letter}\\"]
            
            self.log(f"COMMAND: {' '.join(vss_cmd)}")
            vss_proc = subprocess.run(vss_cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
            
            # Log both stdout and stderr for debugging
            if vss_proc.stdout:
                self.log(f"VSS stdout: {vss_proc.stdout}")
            if vss_proc.stderr:
                self.log(f"VSS stderr: {vss_proc.stderr}")
            
            if vss_proc.returncode != 0:
                self.log(f"ERROR: vssadmin failed with return code: {vss_proc.returncode}")
                if vss_proc.stderr:
                    self.log(f"ERROR: vssadmin error: {vss_proc.stderr}")
                else:
                    self.log("ERROR: vssadmin failed but no error message provided")
                
                # Enhanced troubleshooting
                self.log("TROUBLESHOOTING VSS failure:")
                self.log("  1. Check if VSS service is running: sc query VSS")
                self.log("  2. Run as Administrator")
                self.log("  3. Ensure sufficient disk space (need >15% free)")
                self.log("  4. Check if drive supports VSS")
                self.log("  5. Try: vssadmin list providers")
                self.log("  6. Check Windows version compatibility")
                self.log("  7. Ensure no antivirus blocking VSS operations")
                self.log("  8. Some Windows 10/11 versions have VSS restrictions")
                self.log("  9. Try disabling Windows Defender real-time protection temporarily")
                self.log("  10. Check Event Viewer for VSS errors (System log)")
                self.log("")
                self.log("COMMON VSS SOLUTIONS:")
                self.log("  • Restart Windows and try again")
                self.log("  • Free up disk space to >20% available")
                self.log("  • Run: sfc /scannow")
                self.log("  • Run: dism /online /cleanup-image /restorehealth")
                self.log("  • Disable System Restore temporarily")
                self.log("  • Use Direct method as fallback (but less safe)")
                
                # Try to get more diagnostic info
                self.log("INFO: Running VSS diagnostics...")
                diag_cmds = [
                    ["vssadmin", "list", "providers"],
                    ["vssadmin", "list", "volumes"],
                    ["sc", "query", "VSS"],
                    ["sc", "query", "SWPRV"]
                ]
                
                for diag_cmd in diag_cmds:
                    try:
                        diag_proc = subprocess.run(diag_cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
                        self.log(f"DIAG {' '.join(diag_cmd)}: {diag_proc.stdout}")
                    except:
                        pass
                
                return None
            
            # Parse the output to get shadow copy ID
            output = vss_proc.stdout
            self.log("INFO: VSS command completed successfully, parsing shadow copy ID...")
            
            # Look for "Shadow Copy ID: {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}"
            import re
            shadow_id_match = re.search(r'Shadow Copy ID: \{([^}]+)\}', output)
            if shadow_id_match:
                shadow_id = shadow_id_match.group(1)
                self.log(f"SUCCESS: Created shadow copy with ID: {shadow_id}")
                return shadow_id
            else:
                self.log("ERROR: Could not parse shadow copy ID from vssadmin output")
                self.log(f"Full VSS Output:\n{output}")
                return None
                
        except Exception as e:
            self.log(f"ERROR: Failed to create VSS shadow copy: {e}")
            return None

    def get_vss_shadow_path(self, shadow_id):
        """
        Get the path for a VSS shadow copy.
        Prioritizes PowerShell for better reliability and parsing.
        Falls back to vssadmin if PowerShell fails.
        """
        self.log(f"Attempting to get shadow path for ID: {shadow_id}")

        # Method 1: PowerShell using WMI (most reliable)
        try:
            self.log("Trying to get shadow path using PowerShell (WMI)...")
            # This is more compatible across Windows versions than Get-VssShadow
            ps_command = f"(Get-WmiObject Win32_ShadowCopy -Filter \"ID='{shadow_id}'\").DeviceObject"
            
            result = subprocess.run(
                ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_command],
                capture_output=True, text=True, check=True, timeout=30, encoding='utf-8', errors='ignore'
            )
            
            shadow_path = result.stdout.strip()
            if shadow_path and 'HarddiskVolumeShadowCopy' in shadow_path:
                self.log(f"PowerShell successful. Found shadow path: {shadow_path}")
                return shadow_path
            else:
                self.log(f"PowerShell command ran but output was not a valid path: '{shadow_path}'")

        except Exception as e:
            self.log(f"PowerShell WMI method failed: {e}. Falling back to vssadmin.")

        # Method 2: vssadmin (fallback with robust parsing)
        self.log("Trying to get shadow path using vssadmin...")
        for attempt in range(1, 6):
            self.log(f"vssadmin attempt {attempt}/5...")
            try:
                result = subprocess.run(
                    ["vssadmin", "list", "shadows", f"/shadow={shadow_id}"],
                    capture_output=True, text=True, check=True, timeout=60, encoding='utf-8', errors='ignore'
                )
                output = result.stdout
                
                # Use robust string splitting instead of regex
                for line in output.splitlines():
                    if "Shadow Copy Volume:" in line:
                        path = line.split(":", 1)[1].strip()
                        if "HarddiskVolumeShadowCopy" in path:
                            self.log(f"vssadmin successful. Found shadow path: {path}")
                            return path

                self.log(f"Could not parse shadow path from vssadmin output on attempt {attempt}.")
                self.log(f"Full vssadmin output:\n{output}")

            except Exception as e:
                self.log(f"An unexpected error occurred with vssadmin on attempt {attempt}: {e}")
            
            time.sleep(5)  # Wait before retrying

        self.log(f"Error: Failed to get shadow path for ID {shadow_id} after all attempts.")
        return None

    def delete_vss_shadow_copy(self, shadow_id):
        """Delete a VSS shadow copy using a more robust PowerShell command."""
        try:
            if not shadow_id:
                self.log("INFO: No shadow ID provided, skipping cleanup.")
                return

            self.log(f"INFO: Deleting VSS shadow copy: {shadow_id}")
            
            # Use PowerShell for more reliable deletion. WMI filter needs the ID with braces.
            ps_cmd = f"(Get-WmiObject Win32_ShadowCopy -Filter \"ID='{shadow_id}'\").Delete()"
            
            ps_proc = subprocess.run([
                "powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps_cmd
            ], capture_output=True, text=True, encoding='utf-8', errors='ignore')

            if ps_proc.returncode == 0 and not ps_proc.stderr:
                self.log(f"SUCCESS: VSS shadow copy {shadow_id} deleted successfully.")
            else:
                self.log(f"WARNING: Failed to delete VSS shadow copy {shadow_id}.")
                self.log(f"  PowerShell stderr: {ps_proc.stderr.strip()}")
                self.log(f"  PowerShell stdout: {ps_proc.stdout.strip()}")
                self.log("  Fallback: Trying vssadmin...")
                
                vss_cmd = ["vssadmin", "delete", "shadows", f"/shadow={shadow_id}", "/quiet"]
                vss_proc_fallback = subprocess.run(vss_cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
                if vss_proc_fallback.returncode == 0:
                    self.log("SUCCESS: Fallback vssadmin delete succeeded.")
                else:
                    self.log(f"ERROR: Fallback vssadmin delete also failed: {vss_proc_fallback.stderr.strip()}")

        except Exception as e:
            self.log(f"WARNING: Exception during VSS cleanup: {e}")

    def create_vss_drive_mapping(self, shadow_path):
        """Create a temporary drive letter mapping for VSS shadow copy path"""
        try:
            # Find an available drive letter
            available_letters = []
            for letter in "ZYXWVUTSRQPONMLKJIHGFED":  # Start from Z and work backwards
                drive_path = f"{letter}:\\"
                if not os.path.exists(drive_path):
                    available_letters.append(letter)
            
            if not available_letters:
                self.log("WARNING: No available drive letters for VSS mapping")
                return None
                
            drive_letter = available_letters[0]
            self.log(f"INFO: Creating drive mapping {drive_letter}: -> {shadow_path}")
            
            # Method 1: Try subst command with proper escaping
            # Remove any trailing backslashes and requote properly
            clean_shadow_path = shadow_path.rstrip('\\')
            subst_cmd = ['subst', f'{drive_letter}:', clean_shadow_path]
            self.log(f"DEBUG: Running command: {' '.join(subst_cmd)}")
            result = subprocess.run(subst_cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
            
            if result.returncode == 0:
                # Give Windows a moment to register the mapping
                time.sleep(1)
                # Verify the mapping worked
                test_drive = f"{drive_letter}:\\"
                if os.path.exists(test_drive):
                    self.log(f"SUCCESS: Created drive mapping {drive_letter}: and verified accessible")
                    return test_drive
                else:
                    self.log(f"ERROR: Drive mapping created but not accessible: {test_drive}")
                    # Try to see what drives exist
                    try:
                        import win32api
                        drives = win32api.GetLogicalDriveStrings()
                        self.log(f"DEBUG: Available drives: {drives}")
                    except:
                        pass
            else:
                self.log(f"ERROR: subst command failed with return code {result.returncode}")
                self.log(f"ERROR: subst stderr: {result.stderr.strip()}")
                self.log(f"ERROR: subst stdout: {result.stdout.strip()}")
                
            # Method 1b: Try subst with PowerShell (sometimes works better)
            self.log("INFO: Trying subst via PowerShell...")
            ps_subst_cmd = f'subst {drive_letter}: "{clean_shadow_path}"'
            ps_result = subprocess.run(['powershell', '-Command', ps_subst_cmd], 
                                     capture_output=True, text=True, encoding='utf-8', errors='ignore')
            
            if ps_result.returncode == 0:
                time.sleep(1)
                test_drive = f"{drive_letter}:\\"
                if os.path.exists(test_drive):
                    self.log(f"SUCCESS: PowerShell subst created drive mapping {drive_letter}:")
                    return test_drive
                else:
                    self.log(f"ERROR: PowerShell subst mapping created but not accessible: {test_drive}")
            else:
                self.log(f"ERROR: PowerShell subst failed: {ps_result.stderr.strip()}")
            
            # Method 2: Try creating directory junction (often works better with VSS)
            self.log("INFO: Trying alternative method - creating directory junction...")
            try:
                # Create a temporary directory to use as a mount point
                temp_dir = f"C:\\temp_vss_mount_{drive_letter}"
                
                # Remove existing temp dir if it exists
                if os.path.exists(temp_dir):
                    try:
                        subprocess.run(['rmdir', '/S', '/Q', temp_dir], shell=True, capture_output=True)
                    except:
                        pass
                
                os.makedirs(temp_dir, exist_ok=True)
                
                # Try mklink to create directory junction with clean path
                mklink_cmd = ['mklink', '/J', temp_dir, clean_shadow_path]
                self.log(f"DEBUG: Running mklink command: {' '.join(mklink_cmd)}")
                mklink_result = subprocess.run(mklink_cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
                
                if mklink_result.returncode == 0:
                    # Give Windows a moment and verify
                    time.sleep(1)
                    if os.path.exists(temp_dir) and os.path.isdir(temp_dir):
                        try:
                            # Test if we can list contents
                            test_files = os.listdir(temp_dir)
                            self.log(f"SUCCESS: Created working junction at {temp_dir} (found {len(test_files)} items)")
                            return temp_dir + "\\"
                        except Exception as e:
                            self.log(f"ERROR: Junction created but not accessible: {e}")
                    else:
                        self.log(f"ERROR: Junction directory doesn't exist after creation")
                else:
                    self.log(f"ERROR: mklink failed with code {mklink_result.returncode}")
                    self.log(f"ERROR: mklink stderr: {mklink_result.stderr.strip()}")
                    self.log(f"ERROR: mklink stdout: {mklink_result.stdout.strip()}")
                    
            except Exception as e:
                self.log(f"ERROR: Junction creation failed: {e}")
            
            return None
                
        except Exception as e:
            self.log(f"ERROR: Exception creating VSS drive mapping: {e}")
            return None
    
    def remove_vss_drive_mapping(self, mapped_path):
        """Remove temporary drive letter mapping or junction"""
        try:
            if not mapped_path:
                return
                
            self.log(f"INFO: Cleaning up VSS mapping: {mapped_path}")
            
            # Check if it's a drive letter mapping (e.g., "Z:\")
            if len(mapped_path) >= 3 and mapped_path[1:3] == ':\\':
                letter = mapped_path[0].upper()
                self.log(f"INFO: Removing drive mapping {letter}:")
                
                subst_cmd = f'subst {letter}: /D'
                result = subprocess.run(subst_cmd, shell=True, capture_output=True, text=True, encoding='utf-8', errors='ignore')
                
                if result.returncode == 0:
                    self.log(f"SUCCESS: Removed drive mapping {letter}:")
                else:
                    self.log(f"WARNING: Failed to remove drive mapping: {result.stderr}")
            
            # Check if it's a junction/temp directory (e.g., "C:\temp_vss_mount_Z\")
            elif mapped_path.startswith("C:\\temp_vss_mount_"):
                self.log(f"INFO: Removing junction directory: {mapped_path}")
                try:
                    # Remove junction using rmdir
                    rmdir_cmd = f'rmdir "{mapped_path.rstrip(chr(92))}"'
                    result = subprocess.run(rmdir_cmd, shell=True, capture_output=True, text=True, encoding='utf-8', errors='ignore')
                    
                    if result.returncode == 0:
                        self.log(f"SUCCESS: Removed junction directory")
                    else:
                        self.log(f"WARNING: Failed to remove junction: {result.stderr}")
                        # Try with /S flag to force removal
                        rmdir_cmd_force = f'rmdir /S /Q "{mapped_path.rstrip(chr(92))}"'
                        result2 = subprocess.run(rmdir_cmd_force, shell=True, capture_output=True, text=True, encoding='utf-8', errors='ignore')
                        if result2.returncode == 0:
                            self.log(f"SUCCESS: Force removed junction directory")
                        else:
                            self.log(f"WARNING: Force removal also failed: {result2.stderr}")
                            
                except Exception as e:
                    self.log(f"WARNING: Exception removing junction: {e}")
                    
        except Exception as e:
            self.log(f"WARNING: Exception removing VSS mapping: {e}")

    def create_direct_wim_image(self):
        """Creates a WIM image using DISM directly (risky approach)."""
        # This method is deprecated - use the modern restic backup workflow instead
        self.log("ERROR: This DISM method is deprecated. Use the modern Restic backup in Step 1.")
        messagebox.showerror("Method Deprecated", 
                           "DISM-based WIM creation has been replaced with the modern Restic workflow in Step 1.\n\n" +
                           "Please use 'Create System Backup' in Step 1 instead.")
        return False

    def start_vss_wim_creation_thread(self):
        """Starts the VSS + DISM WIM creation process in a new thread."""
        # Disable buttons
        if hasattr(self, 'vss_create_button'):
            self.vss_create_button.config(state="disabled")
        if hasattr(self, 'direct_create_button'):
            self.direct_create_button.config(state="disabled")
        
        thread = threading.Thread(target=self.vss_wim_creation_worker)
        thread.daemon = True
        thread.start()

    def start_direct_wim_creation_thread(self):
        """Starts the direct DISM WIM creation process in a new thread."""
        # Disable buttons
        if hasattr(self, 'vss_create_button'):
            self.vss_create_button.config(state="disabled")
        if hasattr(self, 'direct_create_button'):
            self.direct_create_button.config(state="disabled")
        
        thread = threading.Thread(target=self.direct_wim_creation_worker)
        thread.daemon = True
        thread.start()

    def vss_wim_creation_worker(self):
        """Worker function for VSS + DISM WIM creation."""
        try:
            success = self.create_vss_wim_image()
            if success:
                self.log("=== VSS + DISM WIM CREATION COMPLETED SUCCESSFULLY ===")
            else:
                self.log("=== VSS + DISM WIM CREATION FAILED ===")
        except Exception as e:
            self.log(f"FATAL: VSS + DISM worker failed: {e}")
        finally:
            # Re-enable buttons
            if hasattr(self, 'vss_create_button'):
                self.vss_create_button.config(state="normal")
            if hasattr(self, 'direct_create_button'):
                self.direct_create_button.config(state="normal")

    def direct_wim_creation_worker(self):
        """Worker function for direct DISM WIM creation (risky)."""
        try:
            success = self.create_direct_wim_image()
            if success:
                self.log("=== DIRECT WIM CREATION COMPLETED ===")
            else:
                self.log("=== DIRECT WIM CREATION FAILED ===")
        except Exception as e:
            self.log(f"FATAL: Direct WIM creation failed: {e}")
        finally:
            # Re-enable buttons
            if hasattr(self, 'vss_create_button'):
                self.vss_create_button.config(state="normal")
            if hasattr(self, 'direct_create_button'):
                self.direct_create_button.config(state="normal")

    def download_restic(self):
        """Download the Restic v0.18.0 binary for Windows."""
        try:
            restic_dir = Path("./restic")
            restic_dir.mkdir(exist_ok=True)
            restic_exe = restic_dir / "restic.exe"
            
            if restic_exe.exists():
                # Check if it's a valid executable
                try:
                    version_proc = subprocess.run([str(restic_exe), "version"], 
                                                capture_output=True, text=True, timeout=10)
                    if version_proc.returncode == 0:
                        self.log(f"INFO: Restic binary already exists: {version_proc.stdout.strip()}")
                        return str(restic_exe)
                    else:
                        self.log("WARNING: Existing restic.exe appears corrupted, re-downloading...")
                        restic_exe.unlink()
                except Exception:
                    self.log("WARNING: Could not verify existing restic.exe, re-downloading...")
                    restic_exe.unlink()
            
            self.log("INFO: Downloading Restic v0.18.0 binary for Windows...")
            
            # Use the direct download URL for Restic v0.18.0
            download_url = "https://github.com/restic/restic/releases/download/v0.18.0/restic_0.18.0_windows_amd64.zip"
            self.log(f"INFO: Downloading from: {download_url}")
            
            # Download and extract
            import zipfile
            import tempfile
            import requests
            
            # Create temporary file for download
            tmp_file_path = None
            try:
                with tempfile.NamedTemporaryFile(suffix='.zip', delete=False) as tmp_file:
                    tmp_file_path = tmp_file.name
                    self.log("INFO: Starting download...")
                    zip_response = requests.get(download_url, stream=True, timeout=120)
                    zip_response.raise_for_status()
                    
                    total_size = int(zip_response.headers.get('content-length', 0))
                    downloaded = 0
                    last_progress_logged = -1
                    
                    for chunk in zip_response.iter_content(chunk_size=8192):
                        if chunk:
                            tmp_file.write(chunk)
                            downloaded += len(chunk)
                            if total_size > 0:
                                progress = (downloaded / total_size) * 100
                                # Log progress every 10%
                                if int(progress / 10) > last_progress_logged:
                                    last_progress_logged = int(progress / 10)
                                    self.log(f"INFO: Download progress: {progress:.1f}%")
                    
                    tmp_file.flush()
                
                self.log("INFO: Download completed, extracting...")
                
                # Add a small delay to ensure file handle is fully released
                time.sleep(0.5)
                
                # Extract restic.exe from zip - the file should be named restic_0.18.0_windows_amd64.exe
                try:
                    with zipfile.ZipFile(tmp_file_path, 'r') as zip_ref:
                        extracted = False
                        for file_info in zip_ref.filelist:
                            # Look for the specific executable name or any .exe file
                            if (file_info.filename == 'restic_0.18.0_windows_amd64.exe' or 
                                file_info.filename.endswith('.exe')):
                                self.log(f"INFO: Extracting {file_info.filename}...")
                                # Extract to our restic directory
                                with zip_ref.open(file_info) as source, open(restic_exe, 'wb') as target:
                                    target.write(source.read())
                                extracted = True
                                break
                        
                        if not extracted:
                            self.log("ERROR: Could not find restic.exe in the downloaded ZIP file")
                            self.log(f"ZIP contents: {[f.filename for f in zip_ref.filelist]}")
                            return None
                    
                    self.log("INFO: Extraction completed")
                    
                except zipfile.BadZipFile as e:
                    self.log(f"ERROR: Downloaded file is not a valid ZIP archive: {e}")
                    return None
                    
            finally:
                # Clean up temp file with retry logic
                if tmp_file_path and Path(tmp_file_path).exists():
                    for attempt in range(3):
                        try:
                            Path(tmp_file_path).unlink()
                            self.log("INFO: Temporary download file cleaned up")
                            break
                        except PermissionError:
                            if attempt < 2:
                                self.log(f"INFO: Retrying temp file cleanup (attempt {attempt + 1}/3)")
                                time.sleep(1)
                            else:
                                self.log(f"WARNING: Could not clean up temporary file: {tmp_file_path}")
            
            if restic_exe.exists():
                # Verify the downloaded binary works
                try:
                    self.log("INFO: Verifying downloaded Restic binary...")
                    version_proc = subprocess.run([str(restic_exe), "version"], 
                                                capture_output=True, text=True, timeout=10)
                    if version_proc.returncode == 0:
                        version_info = version_proc.stdout.strip()
                        self.log(f"SUCCESS: Restic downloaded and verified: {version_info}")
                        # Check if it's the expected version
                        if "0.18.0" in version_info:
                            self.log("SUCCESS: Confirmed Restic v0.18.0")
                        else:
                            self.log(f"WARNING: Downloaded version might not be v0.18.0: {version_info}")
                        return str(restic_exe)
                    else:
                        self.log(f"ERROR: Downloaded Restic binary failed verification: {version_proc.stderr}")
                        restic_exe.unlink()  # Remove corrupted file
                        return None
                except Exception as e:
                    self.log(f"ERROR: Could not verify downloaded Restic binary: {e}")
                    return None
            else:
                self.log("ERROR: Failed to extract restic.exe from download")
                return None
                
        except Exception as e:
            self.log(f"ERROR: Failed to download Restic: {e}")
            return None

    def create_vss_restic_backup(self):
        """Creates a backup using Restic's built-in VSS support."""
        try:
            self.log("INFO: Starting VSS + Restic backup process...")
            
            # Validate configuration
            if not self.validate_backup_config():
                return False
            
            # Download Restic if needed
            restic_exe = self.download_restic()
            if not restic_exe:
                self.log("ERROR: Failed to get Restic binary")
                return False
            
            # Initialize repository if needed
            if not self.init_restic_repository():
                return False
            
            # Perform the backup
            return self.perform_restic_backup(restic_exe)
            
        except Exception as e:
            self.log(f"FATAL: VSS + Restic backup failed: {e}")
            return False

    def validate_backup_config(self):
        """Validate backup configuration"""
        try:
            # Check if we have required configuration
            repo_type = getattr(self, 'repo_type_var', None)
            if not repo_type:
                self.log("ERROR: Repository type not configured")
                return False
            
            self.log(f"INFO: Repository type selected: {repo_type.get()}")
                
            if repo_type.get() == "s3":
                # Validate S3 configuration
                s3_config = self.get_s3_config_for_mode()
                if not s3_config or not all([s3_config.get('s3_bucket'), s3_config.get('s3_access_key'), s3_config.get('s3_secret_key')]):
                    self.log("ERROR: S3 configuration incomplete. Please configure S3 settings.")
                    self.log("SOLUTION: Click 'Configure S3...' button to set up S3 credentials")
                    self.log("ALTERNATIVE: Switch to 'Local File System' repository type")
                    
                    # Debug: Show what S3 config we found
                    if s3_config:
                        self.log(f"DEBUG: Found S3 config keys: {list(s3_config.keys())}")
                    else:
                        self.log("DEBUG: No S3 config found in database")
                    
                    # Show helpful dialog
                    messagebox.showerror("S3 Configuration Required", 
                        "S3 Cloud Storage is selected but not configured.\n\n" +
                        "Please either:\n" +
                        "1. Click 'Configure S3...' button to set up S3 credentials\n" +
                        "2. Switch to 'Local File System' repository type\n\n" +
                        "Then try the backup again.")
                    return False
                    
            elif repo_type.get() == "local":
                # Validate local path
                repo_location = getattr(self, 'repo_location_var', None)
                if not repo_location or not repo_location.get():
                    self.log("ERROR: Local repository path not configured")
                    self.log("SOLUTION: Set a local folder path for the repository")
                    
                    # Show helpful dialog
                    messagebox.showerror("Local Path Required", 
                        "Local File System is selected but no path is configured.\n\n" +
                        "Please set a local folder path where the backup repository will be stored.\n\n" +
                        "Example: C:\\Backups\\ResticRepo")
                    return False
                    
            return True
        except Exception as e:
            self.log(f"ERROR: Configuration validation failed: {e}")
            return False

    def init_restic_repository(self):
        """Initialize restic repository if needed"""
        try:
            self.log("INFO: Checking/initializing Restic repository...")
            
            # Get restic executable
            restic_exe = self.download_restic()
            if not restic_exe:
                return False
            
            # Set up environment
            env = os.environ.copy()
            
            # Get or generate secure repository password
            repo_password = self.get_or_generate_repository_password()
            if not repo_password:
                self.log("ERROR: Failed to get or generate repository password")
                return False
            
            env['RESTIC_PASSWORD'] = repo_password
            
            repo_type = self.repo_type_var.get() if hasattr(self, 'repo_type_var') else "s3"
            if repo_type == "s3":
                s3_config = self.get_s3_config_for_mode()
                
                # Build organized S3 path structure
                s3_repo_path = self.build_s3_repository_path(s3_config)
                if not s3_repo_path:
                    return False
                    
                env['RESTIC_REPOSITORY'] = s3_repo_path
                if s3_config and s3_config.get('s3_access_key'):
                    env['AWS_ACCESS_KEY_ID'] = s3_config['s3_access_key']
                if s3_config and s3_config.get('s3_secret_key'):
                    env['AWS_SECRET_ACCESS_KEY'] = s3_config['s3_secret_key']
            else:
                repo_path = self.repo_location_var.get()
                env['RESTIC_REPOSITORY'] = repo_path
                
                # Create local directory if it doesn't exist
                Path(repo_path).mkdir(parents=True, exist_ok=True)
            
            # Try to initialize repository first to determine if it exists
            self.log("INFO: Checking repository status...")
            
            # Try to initialize - this will fail if repository already exists
            init_cmd = [str(restic_exe), "init"]
            init_proc = subprocess.run(
                init_cmd,
                env=env,
                capture_output=True,
                text=True
            )
            
            if init_proc.returncode == 0:
                # Repository was successfully initialized (was new)
                self.log("SUCCESS: New restic repository initialized successfully")
                return True
            else:
                # Check if failure was due to existing repository
                if "already initialized" in init_proc.stderr:
                    self.log("INFO: Repository already exists - requesting password confirmation...")
                    
                    # Prompt user for existing repository password
                    confirmed_password = self.prompt_repository_password_confirmation()
                    if not confirmed_password:
                        self.log("ERROR: Password confirmation cancelled or failed")
                        return False
                    
                    # Update environment with confirmed password
                    env['RESTIC_PASSWORD'] = confirmed_password
                    
                    # Verify password works
                    verify_cmd = [str(restic_exe), "snapshots", "--json", "--last", "1"]
                    verify_proc = subprocess.run(
                        verify_cmd,
                        env=env,
                        capture_output=True,
                        text=True
                    )
                    
                    if verify_proc.returncode == 0:
                        self.log("SUCCESS: Repository password verified successfully")
                        return True
                    else:
                        self.log("ERROR: Repository password verification failed - incorrect password")
                        return False
                else:
                    # Some other initialization error
                    self.log(f"ERROR: Repository initialization failed: {init_proc.stderr}")
                    return False
                    
        except Exception as e:
            self.log(f"ERROR: Repository initialization failed: {e}")
            return False

    def get_s3_config_for_mode(self):
        """Get S3 configuration based on current workflow mode"""
        workflow_mode = self.get_workflow_mode()
        
        if workflow_mode == "development":
            # Development mode: use UI variables
            return {
                "s3_bucket": self.dev_s3_bucket_var.get(),
                "s3_access_key": self.dev_s3_access_var.get(),
                "s3_secret_key": self.dev_s3_secret_var.get(),
                "s3_endpoint": self.dev_s3_endpoint_var.get(),
                "s3_region": self.dev_s3_region_var.get()
            }
        else:
            # Production mode: use database
            return self.db.get_s3_config()

    def get_or_generate_repository_password(self):
        """Get existing repository password or generate a new secure one"""
        try:
            # Get current client/site information for password identification
            workflow_mode = self.get_workflow_mode()
            client_uuid = None
            site_uuid = None
            client_name = "unknown"
            site_name = "unknown"
            
            if workflow_mode == "development":
                # Development mode: use dev_client_var and S3 metadata
                if hasattr(self, 'dev_client_var') and self.dev_client_var.get():
                    client_short = self.dev_client_var.get().split(' (')[0]
                    if hasattr(self, 's3_clients'):
                        for uuid, data in self.s3_clients.items():
                            if data['short_name'] == client_short:
                                client_uuid = uuid
                                client_name = data['name']
                                break
                
                if hasattr(self, 'dev_site_var') and self.dev_site_var.get():
                    site_short = self.dev_site_var.get().split(' (')[0]
                    if hasattr(self, 's3_sites'):
                        for uuid, data in self.s3_sites.items():
                            if data['short_name'] == site_short:
                                site_uuid = uuid
                                site_name = data['name']
                                break
            else:
                # Production mode: use client_var and database
                client_var = getattr(self, 'client_var', None)
                if client_var and client_var.get() and client_var.get() != "-- Select Client --":
                    try:
                        clients = self.db.get_clients()
                        for cid, name, short_name, desc in clients:
                            if name == client_var.get():
                                client_uuid = cid
                                client_name = name
                                break
                    except Exception as e:
                        self.log(f"WARNING: Could not retrieve client info: {e}")
                
                site_var = getattr(self, 'site_var', None)
                if site_var and site_var.get() and site_var.get() != "-- Select Site --":
                    try:
                        sites = self.db.get_sites()
                        for sid, client_id, name, short_name, desc, client_name_db in sites:
                            if name == site_var.get():
                                site_uuid = sid
                                site_name = name
                                break
                    except Exception as e:
                        self.log(f"WARNING: Could not retrieve site info: {e}")
            
            # Create a unique identifier for this repository
            repo_identifier = f"{client_uuid or 'default-client'}_{site_uuid or 'default-site'}_backup"
            
            # Try to get existing password from database config
            existing_password = self.db.get_config(f"restic_password_{repo_identifier}")
            
            if existing_password:
                self.log("INFO: Using existing repository password")
                return existing_password
            
            # Generate new secure password
            self.log("INFO: Generating new secure repository password...")
            role = getattr(self, 'role_var', None)
            role_name = role.get() if role else "backup"
            
            new_password, password_identifier = self.db.generate_secure_password(client_name, site_name, role_name)
            
            # Store the password in database for future use
            self.db.set_config(f"restic_password_{repo_identifier}", new_password)
            
            # Display password to user for record keeping
            self.show_repository_password_reminder(new_password, password_identifier, client_name, site_name, role_name)
            
            self.log("SUCCESS: Repository password generated and stored")
            return new_password
            
        except Exception as e:
            self.log(f"ERROR: Failed to get or generate repository password: {e}")
            return None

    def show_repository_password_reminder(self, password, identifier, client_name="", site_name="", role=""):
        """Show dialog to remind user to save the repository password"""
        try:
            dialog = self.create_centered_dialog("🔐 Repository Password - SAVE THIS!", 700, 500)
            
            # Warning frame
            warning_frame = ttk.LabelFrame(dialog, text="⚠️ CRITICAL - Repository Password", padding="15")
            warning_frame.pack(fill="x", padx=20, pady=(20, 10))
            
            warning_text = """This is your RESTIC REPOSITORY PASSWORD. You MUST save this password!

• Without this password, you CANNOT access your backups
• This password is required for ALL restore operations
• Store this password in your password manager
• Make a backup copy in a secure location

LOSING THIS PASSWORD = LOSING ALL YOUR BACKUPS!"""
            
            ttk.Label(warning_frame, text=warning_text, 
                     font=("TkDefaultFont", 10, "bold"), 
                     foreground="red", justify="left").pack()
            
            # Repository info frame
            info_frame = ttk.LabelFrame(dialog, text="Repository Information", padding="15")
            info_frame.pack(fill="x", padx=20, pady=10)
            info_frame.columnconfigure(1, weight=1)
            
            ttk.Label(info_frame, text="Client:", font=("TkDefaultFont", 9, "bold")).grid(row=0, column=0, sticky="w", pady=2)
            ttk.Label(info_frame, text=client_name).grid(row=0, column=1, sticky="w", padx=(10, 0), pady=2)
            
            ttk.Label(info_frame, text="Site:", font=("TkDefaultFont", 9, "bold")).grid(row=1, column=0, sticky="w", pady=2)
            ttk.Label(info_frame, text=site_name).grid(row=1, column=1, sticky="w", padx=(10, 0), pady=2)
            
            ttk.Label(info_frame, text="Role:", font=("TkDefaultFont", 9, "bold")).grid(row=2, column=0, sticky="w", pady=2)
            ttk.Label(info_frame, text=role).grid(row=2, column=1, sticky="w", padx=(10, 0), pady=2)
            
            ttk.Label(info_frame, text="Repository ID:", font=("TkDefaultFont", 9, "bold")).grid(row=3, column=0, sticky="w", pady=2)
            ttk.Label(info_frame, text=identifier).grid(row=3, column=1, sticky="w", padx=(10, 0), pady=2)
            
            # Password frame
            password_frame = ttk.LabelFrame(dialog, text="Repository Password", padding="15")
            password_frame.pack(fill="x", padx=20, pady=10)
            
            # Password display with copy button
            password_display_frame = ttk.Frame(password_frame)
            password_display_frame.pack(fill="x", pady=5)
            
            password_entry = ttk.Entry(password_display_frame, font=("Consolas", 12), width=50)
            password_entry.pack(side="left", fill="x", expand=True)
            password_entry.insert(0, password)
            password_entry.config(state="readonly")
            
            def copy_password():
                dialog.clipboard_clear()
                dialog.clipboard_append(password)
                copy_btn.config(text="✓ Copied!")
                dialog.after(2000, lambda: copy_btn.config(text="📋 Copy"))
            
            copy_btn = ttk.Button(password_display_frame, text="📋 Copy", command=copy_password, width=10)
            copy_btn.pack(side="right", padx=(10, 0))
            
            # Instructions
            instructions_frame = ttk.LabelFrame(dialog, text="Next Steps", padding="15")
            instructions_frame.pack(fill="x", padx=20, pady=10)
            
            instructions = """1. IMMEDIATELY copy this password to your password manager
2. Create a backup copy and store it securely
3. Test that you can access the password when needed
4. Click 'I Have Saved The Password' only after completing steps 1-3"""
            
            ttk.Label(instructions_frame, text=instructions, justify="left").pack()
            
            # Button frame
            button_frame = ttk.Frame(dialog)
            button_frame.pack(fill="x", padx=20, pady=20)
            
            saved_var = tk.BooleanVar()
            
            def acknowledge():
                if saved_var.get():
                    dialog.destroy()
                else:
                    messagebox.showwarning("Password Not Saved", 
                                         "Please check the box to confirm you have saved the password!")
            
            ttk.Checkbutton(button_frame, text="✅ I have saved this password in a secure location", 
                           variable=saved_var).pack(pady=10)
            
            ttk.Button(button_frame, text="I Have Saved The Password", 
                      command=acknowledge, style="Accent.TButton").pack(pady=10)
            
            # Make dialog modal
            dialog.transient(self.root)
            dialog.grab_set()
            dialog.focus_set()
            
        except Exception as e:
            self.log(f"ERROR: Failed to show password reminder: {e}")

    def prompt_repository_password_confirmation(self):
        """Prompt user to enter repository password for existing repositories"""
        try:
            # Get repository information for context
            workflow_mode = self.get_workflow_mode()
            client_name = "Unknown"
            site_name = "Unknown"
            repo_identifier = "unknown"
            
            if workflow_mode == "development":
                # Development mode: get info from S3 metadata
                if hasattr(self, 'dev_client_var') and self.dev_client_var.get():
                    client_short = self.dev_client_var.get().split(' (')[0]
                    if hasattr(self, 's3_clients'):
                        for uuid, data in self.s3_clients.items():
                            if data['short_name'] == client_short:
                                client_name = data['name']
                                break
                
                if hasattr(self, 'dev_site_var') and self.dev_site_var.get():
                    site_short = self.dev_site_var.get().split(' (')[0]
                    if hasattr(self, 's3_sites'):
                        for uuid, data in self.s3_sites.items():
                            if data['short_name'] == site_short:
                                site_name = data['name']
                                break
            
            dialog = self.create_centered_dialog("🔐 Repository Password Required", 600, 400)
            
            # Main instruction frame
            instruction_frame = ttk.LabelFrame(dialog, text="Repository Access", padding="15")
            instruction_frame.pack(fill="x", padx=20, pady=(20, 10))
            
            instruction_text = f"""An existing restic repository was found for:

Client: {client_name}
Site: {site_name}

Please enter the repository password to continue with the backup.
This password was shown when the repository was first created."""
            
            ttk.Label(instruction_frame, text=instruction_text, justify="left").pack()
            
            # Password input frame
            password_frame = ttk.LabelFrame(dialog, text="Enter Repository Password", padding="15")
            password_frame.pack(fill="x", padx=20, pady=10)
            
            ttk.Label(password_frame, text="Repository Password:").pack(anchor="w", pady=(0, 5))
            
            password_var = tk.StringVar()
            password_entry = ttk.Entry(password_frame, textvariable=password_var, show="*", width=40, font=("Consolas", 10))
            password_entry.pack(fill="x", pady=(0, 10))
            password_entry.focus_set()
            
            # Show/hide password option
            show_password_var = tk.BooleanVar()
            def toggle_password_visibility():
                if show_password_var.get():
                    password_entry.config(show="")
                else:
                    password_entry.config(show="*")
            
            ttk.Checkbutton(password_frame, text="Show password", 
                           variable=show_password_var, 
                           command=toggle_password_visibility).pack(anchor="w")
            
            # Hint frame
            hint_frame = ttk.LabelFrame(dialog, text="Password Help", padding="10")
            hint_frame.pack(fill="x", padx=20, pady=10)
            
            hint_text = """💡 Hint: The repository password was displayed when this repository was first created.
Check your password manager or secure notes where you saved it."""
            
            ttk.Label(hint_frame, text=hint_text, justify="left", foreground="blue").pack()
            
            # Result variables
            result = {"password": None, "cancelled": False}
            
            # Button frame
            button_frame = ttk.Frame(dialog)
            button_frame.pack(fill="x", padx=20, pady=20)
            
            def confirm_password():
                password = password_var.get().strip()
                if not password:
                    messagebox.showwarning("Password Required", "Please enter the repository password.")
                    password_entry.focus_set()
                    return
                
                result["password"] = password
                dialog.destroy()
            
            def cancel_password():
                result["cancelled"] = True
                dialog.destroy()
            
            # Buttons
            button_container = ttk.Frame(button_frame)
            button_container.pack()
            
            ttk.Button(button_container, text="Cancel", command=cancel_password).pack(side="left", padx=(0, 10))
            ttk.Button(button_container, text="Continue with Backup", command=confirm_password, 
                      style="Accent.TButton").pack(side="left")
            
            # Enter key binding
            password_entry.bind('<Return>', lambda e: confirm_password())
            
            # Make dialog modal
            dialog.transient(self.root)
            dialog.grab_set()
            
            # Wait for dialog to close
            dialog.wait_window()
            
            if result["cancelled"]:
                return None
            
            return result["password"]
            
        except Exception as e:
            self.log(f"ERROR: Failed to prompt for repository password: {e}")
            return None

    def build_s3_repository_path(self, s3_config):
        """Build organized S3 repository path: bucket/client-uuid"""
        try:
            # Validate S3 config
            if not s3_config:
                self.log("ERROR: S3 configuration is None")
                return None
            
            if not isinstance(s3_config, dict):
                self.log("ERROR: S3 configuration is not a dictionary")
                return None
            
            required_keys = ['s3_endpoint', 's3_bucket']
            for key in required_keys:
                if key not in s3_config:
                    self.log(f"ERROR: Missing required S3 config key: {key}")
                    return None
            
            # Get client UUID based on workflow mode
            client_uuid = None
            workflow_mode = self.get_workflow_mode()
            
            if workflow_mode == "development":
                # Development mode: use dev_client_var and S3 metadata
                if hasattr(self, 'dev_client_var') and self.dev_client_var.get():
                    client_short = self.dev_client_var.get().split(' (')[0]
                    if hasattr(self, 's3_clients'):
                        for uuid, data in self.s3_clients.items():
                            if data['short_name'] == client_short:
                                client_uuid = uuid
                                break
            else:
                # Production mode: use client_var and database
                client_name = getattr(self, 'client_var', None)
                if client_name and client_name.get() and client_name.get() != "-- Select Client --":
                    try:
                        clients = self.db.get_clients()
                        for cid, name, short_name, desc in clients:
                            if name == client_name.get():
                                client_uuid = cid
                                break
                    except Exception as e:
                        self.log(f"WARNING: Could not retrieve client UUID: {e}")
            
            # If no client selected, use a default folder
            if not client_uuid:
                client_uuid = "default-client"
                self.log("WARNING: No client selected, using default client folder")
            
            # Build S3 path: s3:endpoint/bucket/client-uuid (no environment subfolder)
            s3_path = f"s3:{s3_config['s3_endpoint']}/{s3_config['s3_bucket']}/{client_uuid}"
            
            self.log(f"INFO: S3 repository structure:")
            self.log(f"  └── Bucket: {s3_config['s3_bucket']}")
            self.log(f"      └── Client: {client_uuid}")
            self.log(f"INFO: Full S3 path: {s3_path}")
            
            return s3_path
            
        except Exception as e:
            self.log(f"ERROR: Failed to build S3 repository path: {e}")
            return None

    def generate_backup_tags(self):
        """Generate comprehensive backup tags with UUIDs and metadata"""
        tags = []
        
        # Generate backup session UUID (UUIDv7)
        backup_uuid = generate_uuidv7()
        tags.append(f"backup-uuid:{backup_uuid}")
        
        # Basic system information
        tags.append("type:system-backup")
        tags.append(f"hostname:{os.environ.get('COMPUTERNAME', 'unknown')}")
        tags.append(f"date:{datetime.now().strftime('%Y-%m-%d')}")
        tags.append(f"time:{datetime.now().strftime('%H-%M-%S')}")
        tags.append(f"timestamp:{int(datetime.now().timestamp())}")
        
        # OS-only vs full backup
        os_only = getattr(self, 'capture_os_only_var', None)
        if os_only and os_only.get():
            tags.append("scope:os-only")
        else:
            tags.append("scope:full-drive")
        
        # Client information
        client_name = getattr(self, 'client_var', None)
        if client_name and client_name.get() and client_name.get() != "-- Select Client --":
            tags.append(f"client-name:{client_name.get()}")
            
            # Find client UUID from database
            try:
                clients = self.db.get_clients()
                for client_id, name, short_name, desc in clients:
                    if name == client_name.get():
                        tags.append(f"client-uuid:{client_id}")
                        tags.append(f"client-short:{short_name}")
                        break
            except Exception as e:
                self.log(f"WARNING: Could not retrieve client UUID: {e}")
        
        # Site information
        site_name = getattr(self, 'site_var', None)
        if site_name and site_name.get() and site_name.get() != "-- Select Site --":
            tags.append(f"site-name:{site_name.get()}")
            
            # Find site UUID from database
            try:
                # Get all sites to find the matching one
                sites = self.db.get_sites()
                for site_id, client_id, name, short_name, desc, client_name_db in sites:
                    if name == site_name.get():
                        tags.append(f"site-uuid:{site_id}")
                        tags.append(f"site-short:{short_name}")
                        break
            except Exception as e:
                self.log(f"WARNING: Could not retrieve site UUID: {e}")
        
        # Role information
        role = getattr(self, 'role_var', None)
        if role and role.get():
            tags.append(f"role:{role.get()}")
        
        # Repository type and configuration
        repo_type = getattr(self, 'repo_type_var', None)
        if repo_type:
            tags.append(f"repo-type:{repo_type.get()}")
            
            if repo_type.get() == "s3":
                try:
                    s3_config = self.get_s3_config_for_mode()
                    if s3_config:
                        tags.append(f"s3-bucket:{s3_config.get('s3_bucket', 'unknown')}")
                        tags.append(f"s3-endpoint:{s3_config.get('s3_endpoint', 'unknown')}")
                except Exception as e:
                    self.log(f"WARNING: Could not retrieve S3 config for tagging: {e}")
        
        # Image type (new vs existing)
        image_type = getattr(self, 'image_type_var', None)
        if image_type and image_type.get():
            tags.append(f"image-type:{image_type.get()}")
        
        # Existing image information if updating
        if image_type and image_type.get() == "existing":
            existing_image = getattr(self, 'existing_image_var', None)
            if existing_image and existing_image.get():
                tags.append(f"base-image:{existing_image.get()}")
        
        # Hardware identification
        hardware_info = self.get_hardware_info()
        if hardware_info['system_uuid']:
            tags.append(f"system-uuid:{hardware_info['system_uuid']}")
        if hardware_info['serial_number']:
            tags.append(f"serial-number:{hardware_info['serial_number']}")
        if hardware_info['manufacturer']:
            tags.append(f"manufacturer:{hardware_info['manufacturer']}")
        if hardware_info['model']:
            tags.append(f"model:{hardware_info['model']}")
        if hardware_info['bios_version']:
            tags.append(f"bios-version:{hardware_info['bios_version']}")
        if hardware_info['total_memory_gb']:
            tags.append(f"memory-gb:{hardware_info['total_memory_gb']}")
        
        # Tool and version information
        tags.append("tool:windows-image-prep-gui")
        tags.append("version:2025.1")  # Tool version
        tags.append(f"restic-version:{self.get_restic_version()}")
        
        # Workflow mode and environment
        workflow_mode = self.get_workflow_mode()
        tags.append(f"workflow-mode:{workflow_mode}")
        
        # Environment tag for organizational purposes
        if workflow_mode == "production":
            environment = "production"
        else:
            environment = "development"  # Default to development
        tags.append(f"environment:{environment}")
        
        # Log hardware summary
        hardware_tags = [tag for tag in tags if any(hw in tag for hw in ['system-uuid', 'serial-number', 'manufacturer', 'model', 'bios-version', 'memory-gb'])]
        if hardware_tags:
            self.log("INFO: Hardware identification tags:")
            for tag in hardware_tags:
                self.log(f"  HW: {tag}")
        
        # Log the tags for debugging
        self.log(f"INFO: Generated {len(tags)} backup tags total")
        # Only log first few tags to avoid spam, hardware tags already logged above
        other_tags = [tag for tag in tags if not any(hw in tag for hw in ['system-uuid', 'serial-number', 'manufacturer', 'model', 'bios-version', 'memory-gb'])]
        self.log(f"INFO: Additional tags: {len(other_tags)} organizational and system tags")
        
        return tags

    def get_restic_version(self):
        """Get Restic version for tagging"""
        try:
            restic_exe = self.download_restic()
            if restic_exe:
                version_proc = subprocess.run([str(restic_exe), "version"], 
                                            capture_output=True, text=True, timeout=10)
                if version_proc.returncode == 0:
                    # Extract version from output like "restic 0.18.0 compiled with go1.24.1 on windows/amd64"
                    version_line = version_proc.stdout.strip()
                    if "restic" in version_line:
                        parts = version_line.split()
                        if len(parts) >= 2:
                            return parts[1]  # Return just the version number
                return "unknown"
        except Exception:
            pass
        return "unknown"

    def get_hardware_info(self):
        """Get hardware information from Windows WMI"""
        hardware_info = {
            'system_uuid': None,
            'serial_number': None,
            'manufacturer': None,
            'model': None,
            'bios_version': None,
            'total_memory_gb': None
        }
        
        try:
            # Use PowerShell to query WMI for hardware information
            # This is more reliable than using Python WMI libraries
            powershell_cmd = '''
            $computer = Get-WmiObject -Class Win32_ComputerSystemProduct
            $bios = Get-WmiObject -Class Win32_BIOS
            $system = Get-WmiObject -Class Win32_ComputerSystem
            
            $output = @{
                "SystemUUID" = $computer.UUID
                "SerialNumber" = $bios.SerialNumber
                "Manufacturer" = $system.Manufacturer
                "Model" = $system.Model
                "BIOSVersion" = $bios.SMBIOSBIOSVersion
                "TotalPhysicalMemory" = [math]::Round($system.TotalPhysicalMemory / 1GB, 2)
            }
            
            $output | ConvertTo-Json -Compress
            '''
            
            self.log("INFO: Retrieving hardware information via WMI...")
            
            result = subprocess.run([
                "powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", powershell_cmd
            ], capture_output=True, text=True, encoding='utf-8', errors='ignore', timeout=30)
            
            if result.returncode == 0 and result.stdout.strip():
                try:
                    import json
                    wmi_data = json.loads(result.stdout.strip())
                    
                    # Extract and clean the data
                    system_uuid = wmi_data.get('SystemUUID', '').strip()
                    serial_number = wmi_data.get('SerialNumber', '').strip()
                    manufacturer = wmi_data.get('Manufacturer', '').strip()
                    model = wmi_data.get('Model', '').strip()
                    bios_version = wmi_data.get('BIOSVersion', '').strip()
                    total_memory = wmi_data.get('TotalPhysicalMemory', 0)
                    
                    # Clean up common placeholder values
                    if system_uuid and system_uuid.lower() not in ['', 'null', 'none', '00000000-0000-0000-0000-000000000000']:
                        hardware_info['system_uuid'] = system_uuid
                        self.log(f"INFO: System UUID: {system_uuid}")
                    
                    if serial_number and serial_number.lower() not in ['', 'null', 'none', 'to be filled by o.e.m.', 'default string']:
                        hardware_info['serial_number'] = serial_number
                        self.log(f"INFO: Serial Number: {serial_number}")
                    
                    if manufacturer and manufacturer.lower() not in ['', 'null', 'none', 'to be filled by o.e.m.']:
                        hardware_info['manufacturer'] = manufacturer.replace(' ', '-').lower()
                        self.log(f"INFO: Manufacturer: {manufacturer}")
                    
                    if model and model.lower() not in ['', 'null', 'none', 'to be filled by o.e.m.']:
                        hardware_info['model'] = model.replace(' ', '-').lower()
                        self.log(f"INFO: Model: {model}")
                    
                    if bios_version and bios_version.lower() not in ['', 'null', 'none']:
                        hardware_info['bios_version'] = bios_version.replace(' ', '-').lower()
                        self.log(f"INFO: BIOS Version: {bios_version}")
                    
                    if total_memory and total_memory > 0:
                        hardware_info['total_memory_gb'] = str(total_memory)
                        self.log(f"INFO: Total Memory: {total_memory} GB")
                    
                except json.JSONDecodeError as e:
                    self.log(f"WARNING: Failed to parse WMI JSON output: {e}")
                except Exception as e:
                    self.log(f"WARNING: Failed to process WMI data: {e}")
            else:
                self.log(f"WARNING: PowerShell WMI query failed: {result.stderr}")
                
        except subprocess.TimeoutExpired:
            self.log("WARNING: WMI query timed out after 30 seconds")
        except Exception as e:
            self.log(f"WARNING: Failed to retrieve hardware information: {e}")
        
        # Fallback: try to get some basic info from environment variables
        if not hardware_info['system_uuid']:
            self.log("INFO: Attempting fallback methods for hardware identification...")
            
            # Try to get computer name as a fallback identifier
            computer_name = os.environ.get('COMPUTERNAME', '')
            if computer_name:
                # Create a deterministic UUID based on computer name
                # This isn't a real system UUID but provides some identification
                import hashlib
                name_hash = hashlib.md5(computer_name.encode()).hexdigest()
                fallback_uuid = f"fallback-{name_hash[:8]}-{name_hash[8:12]}-{name_hash[12:16]}-{name_hash[16:20]}-{name_hash[20:32]}"
                hardware_info['system_uuid'] = fallback_uuid
                self.log(f"INFO: Using fallback UUID based on computer name: {fallback_uuid}")
        
        return hardware_info

    def perform_restic_backup(self, restic_exe):
        """Perform the actual restic backup with VSS"""
        try:
            # Check if user wants OS-only backup
            os_only = getattr(self, 'capture_os_only_var', None)
            if os_only and os_only.get():
                self.log("INFO: Starting Restic backup (OS files only) with Volume Shadow Copy (VSS)...")
            else:
                self.log("INFO: Starting Restic backup (full C: drive) with Volume Shadow Copy (VSS)...")
            
            # Use custom backup tags if provided, otherwise generate them
            if hasattr(self, '_current_backup_tags') and self._current_backup_tags:
                backup_tags = self._current_backup_tags
                self.log("INFO: Using custom backup tags for development mode")
                # Clear the custom tags after use
                delattr(self, '_current_backup_tags')
            else:
                # Generate comprehensive tags with UUIDs and metadata
                backup_tags = self.generate_backup_tags()
            
            # Build restic backup command with explicit VSS support
            backup_cmd = [
                str(restic_exe), "backup", "C:/", 
                "--use-fs-snapshot",  # Enable Volume Shadow Copy on Windows
                "--verbose",
                "--with-atime"  # Include access time (useful for Windows)
            ]
            
            # Add all tags to command
            for tag in backup_tags:
                backup_cmd.extend(["--tag", tag])
            
            # Add exclusions
            exclusions = [
                "C:/Windows/Temp/*", "C:/Windows/Logs/*", "C:/Windows/Prefetch/*",
                "C:/Temp/*", "C:/$Recycle.Bin/*", "C:/System Volume Information/*",
                "C:/pagefile.sys", "C:/hiberfil.sys", "C:/swapfile.sys",
                "*/Temporary Internet Files/*", "*/AppData/Local/Temp/*",
                "*/AppData/Local/Microsoft/Windows/INetCache/*",
                "C:/Windows/SoftwareDistribution/*",  # Windows Update files
                "C:/Windows/Installer/*",  # MSI installer cache
                "C:/ProgramData/Microsoft/Windows/WER/*",  # Windows Error Reporting
                "*/OneDrive*",  # OneDrive cloud-only files that cause VSS issues
                "*OneDrive*",   # Additional OneDrive pattern
            ]
            
            for exclusion in exclusions:
                backup_cmd.extend(["--exclude", exclusion])
            
            self.log(f"INFO: Added {len(exclusions)} standard exclusions for Windows")
            # Note: --one-file-system is not supported on Windows, so we skip it
            self.log("INFO: Skipping --one-file-system flag (not supported on Windows)")
            
            # Add OS-only exclusions if selected
            if os_only and os_only.get():
                self.log("INFO: Adding OS-only exclusions (excluding user data)")
                os_exclusions = [
                    "C:/Users/*/Documents/*", "C:/Users/*/Downloads/*", 
                    "C:/Users/*/Pictures/*", "C:/Users/*/Videos/*",
                    "C:/Users/*/Music/*", "C:/Users/*/Desktop/*",
                    "C:/Users/*/AppData/Local/*", "C:/Users/*/AppData/LocalLow/*"
                ]
                for exclusion in os_exclusions:
                    backup_cmd.extend(["--exclude", exclusion])
                self.log(f"INFO: Added {len(os_exclusions)} OS-only exclusions")
            
            # Set environment variables for repository
            env = os.environ.copy()
            
            repo_type = self.repo_type_var.get() if hasattr(self, 'repo_type_var') else "s3"
            if repo_type == "s3":
                s3_config = self.get_s3_config_for_mode()
                
                # Build organized S3 path structure
                s3_repo_path = self.build_s3_repository_path(s3_config)
                if not s3_repo_path:
                    self.log("ERROR: Failed to build S3 repository path")
                    return False
                    
                env['RESTIC_REPOSITORY'] = s3_repo_path
                if s3_config and s3_config.get('s3_access_key'):
                    env['AWS_ACCESS_KEY_ID'] = s3_config['s3_access_key']
                if s3_config and s3_config.get('s3_secret_key'):
                    env['AWS_SECRET_ACCESS_KEY'] = s3_config['s3_secret_key']
            else:
                env['RESTIC_REPOSITORY'] = self.repo_location_var.get()
            
            # Try to initialize a new repository first to determine if it exists
            self.log("INFO: Checking repository status...")
            
            # Generate new password for potential new repository
            repo_password = self.get_or_generate_repository_password()
            if not repo_password:
                self.log("ERROR: Failed to get or generate repository password")
                return False
            
            env['RESTIC_PASSWORD'] = repo_password
            
            # Try to initialize - this will fail if repository already exists
            init_cmd = [str(restic_exe), "init"]
            init_proc = subprocess.run(
                init_cmd,
                env=env,
                capture_output=True,
                text=True,
                encoding='utf-8',
                errors='ignore'
            )
            
            if init_proc.returncode == 0:
                # Repository was successfully initialized (was new)
                self.log("SUCCESS: New restic repository initialized successfully")
            else:
                # Check if failure was due to existing repository
                if "already initialized" in init_proc.stderr:
                    self.log("INFO: Repository already exists - requesting password confirmation...")
                    
                    # Prompt user for existing repository password
                    confirmed_password = self.prompt_repository_password_confirmation()
                    if not confirmed_password:
                        self.log("ERROR: Password confirmation cancelled or failed")
                        return False
                    
                    # Update environment with confirmed password
                    env['RESTIC_PASSWORD'] = confirmed_password
                    
                    # Verify password works
                    verify_cmd = [str(restic_exe), "snapshots", "--json", "--last", "1"]
                    verify_proc = subprocess.run(
                        verify_cmd,
                        env=env,
                        capture_output=True,
                        text=True
                    )
                    
                    if verify_proc.returncode == 0:
                        self.log("SUCCESS: Repository password verified successfully")
                    else:
                        self.log("ERROR: Repository password verification failed - incorrect password")
                        return False
                else:
                    # Some other initialization error
                    self.log(f"ERROR: Repository initialization failed: {init_proc.stderr}")
                    return False
            
            self.log(f"INFO: Running backup command with {len(backup_cmd)} arguments")
            
            # Run the backup process
            process = subprocess.Popen(
                backup_cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding='utf-8',
                errors='ignore',
                env=env
            )
            
            # Stream output
            while True:
                line = process.stdout.readline()
                if not line:
                    break
                if line.strip():
                    self.log(f"RESTIC: {line.strip()}")
            
            process.wait()
            
            if process.returncode == 0:
                self.log("SUCCESS: Restic backup completed successfully!")
                
                # Store backup metadata in database
                try:
                    self.store_backup_metadata(backup_tags)
                except Exception as e:
                    self.log(f"WARNING: Could not store backup metadata: {e}")
                
                # Store local client metadata JSON file
                try:
                    # Extract client info from tags for metadata
                    tag_dict = {}
                    for tag in backup_tags:
                        if ':' in tag:
                            key, value = tag.split(':', 1)
                            tag_dict[key] = value
                    
                    client_uuid = tag_dict.get('client-uuid')
                    if client_uuid:
                        client_info = self.db_manager.get_client_by_id(client_uuid)
                        site_info = self.db_manager.get_site_by_id(tag_dict.get('site-uuid'))
                        image_info = {
                            "id": tag_dict.get('backup-uuid'),
                            "role": tag_dict.get('role'),
                            "created_at": datetime.now().isoformat()
                        }
                        self.create_client_metadata_json(client_uuid, client_info, site_info, image_info)
                except Exception as e:
                    self.log(f"WARNING: Could not create client metadata JSON: {e}")
                
                # Store metadata JSON file to S3 for discovery
                try:
                    self.store_s3_metadata_file(backup_tags)
                except Exception as e:
                    self.log(f"WARNING: Could not store S3 metadata file: {e}")
                
                return True
            else:
                self.log(f"ERROR: Restic backup failed with exit code: {process.returncode}")
                return False
                
        except Exception as e:
            self.log(f"ERROR: Backup execution failed: {e}")
            return False

    def store_backup_metadata(self, backup_tags):
        """Store backup metadata in database for tracking"""
        try:
            # Extract key information from tags
            backup_uuid = None
            client_uuid = None
            site_uuid = None
            role = None
            scope = None
            repo_type = None
            
            tag_dict = {}
            for tag in backup_tags:
                if ':' in tag:
                    key, value = tag.split(':', 1)
                    tag_dict[key] = value
            
            backup_uuid = tag_dict.get('backup-uuid')
            client_uuid = tag_dict.get('client-uuid') 
            site_uuid = tag_dict.get('site-uuid')
            role = tag_dict.get('role', 'unknown')
            scope = tag_dict.get('scope', 'unknown')
            repo_type = tag_dict.get('repo-type', 'unknown')
            hostname = tag_dict.get('hostname', 'unknown')
            
            if backup_uuid:
                # Store as an image record in the database
                description = f"System backup - {scope} - {hostname} - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
                
                # Use client_uuid if available, otherwise create a generic entry
                client_id = client_uuid if client_uuid else "backup-only"
                site_id = site_uuid if site_uuid else "backup-only"
                
                # Store in images table
                repo_path = "backup-session"  # Special marker for backup sessions
                self.db.create_image(
                    image_id=backup_uuid,
                    client_id=client_id, 
                    site_id=site_id,
                    role=role,
                    repository_path=repo_path,
                    repository_size_gb=0  # Will be updated later if needed
                )
                
                # Store all tags as JSON metadata
                tags_json = json.dumps(tag_dict)
                self.db.save_image_metadata(backup_uuid, tags_json)
                
                self.log(f"INFO: Backup metadata stored with UUID: {backup_uuid}")
            
        except Exception as e:
            self.log(f"ERROR: Failed to store backup metadata: {e}")

    def store_s3_metadata_file(self, backup_tags):
        """Store metadata JSON file to S3 for image discovery"""
        try:
            # Extract backup UUID and other key info
            tag_dict = {}
            for tag in backup_tags:
                if ':' in tag:
                    key, value = tag.split(':', 1)
                    tag_dict[key] = value
            
            backup_uuid = tag_dict.get('backup-uuid')
            if not backup_uuid:
                self.log("WARNING: No backup UUID found, cannot store S3 metadata")
                return False
                
            # Only store metadata for S3 repositories
            repo_type = self.repo_type_var.get() if hasattr(self, 'repo_type_var') else "s3"
            if repo_type != "s3":
                self.log("INFO: Local repository - skipping S3 metadata file")
                return True
                
            # Build comprehensive metadata
            metadata = {
                "backup_uuid": backup_uuid,
                "created_timestamp": datetime.now().isoformat(),
                "version": "1.0",
                "tool": "windows-image-prep-gui",
                "tool_version": "2025.1",
                "tags": tag_dict,
                "hardware": {
                    "system_uuid": tag_dict.get('system-uuid'),
                    "serial_number": tag_dict.get('serial-number'),
                    "manufacturer": tag_dict.get('manufacturer'),
                    "model": tag_dict.get('model'),
                    "hostname": tag_dict.get('hostname'),
                    "memory_gb": tag_dict.get('memory-gb')
                },
                "client_info": {
                    "client_uuid": tag_dict.get('client-uuid'),
                    "client_name": tag_dict.get('client-name'),
                    "site_uuid": tag_dict.get('site-uuid'),
                    "site_name": tag_dict.get('site-name'),
                    "role": tag_dict.get('role')
                },
                "backup_info": {
                    "environment": tag_dict.get('environment'),
                    "scope": tag_dict.get('scope'),
                    "date": tag_dict.get('date'),
                    "time": tag_dict.get('time'),
                    "timestamp": tag_dict.get('timestamp')
                },
                "repository_info": {
                    "repo_type": tag_dict.get('repo-type'),
                    "restic_version": tag_dict.get('restic-version')
                }
            }
            
            # Get S3 configuration
            s3_config = self.db.get_s3_config()
            if not s3_config:
                self.log("ERROR: No S3 configuration found")
                return False
            
            # Build S3 metadata file path in root metadata folder
            # Create JSON content
            json_content = json.dumps(metadata, indent=2, ensure_ascii=False)
            
            # Create temporary file with metadata
            import tempfile
            with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False, encoding='utf-8') as temp_file:
                temp_file.write(json_content)
                temp_file_path = temp_file.name
            
            try:
                # Use AWS CLI to upload metadata file to S3 root metadata folder
                s3_metadata_path = f"s3://{s3_config['s3_bucket']}/metadata/{backup_uuid}.json"
                
                # Set up environment for AWS CLI
                env = os.environ.copy()
                env['AWS_ACCESS_KEY_ID'] = s3_config['s3_access_key']
                env['AWS_SECRET_ACCESS_KEY'] = s3_config['s3_secret_key']
                if s3_config.get('s3_endpoint') and not s3_config['s3_endpoint'].startswith('s3.'):
                    env['AWS_ENDPOINT_URL'] = f"https://{s3_config['s3_endpoint']}"
                
                # Upload using AWS CLI
                aws_cmd = [
                    "aws", "s3", "cp", temp_file_path, s3_metadata_path,
                    "--content-type", "application/json"
                ]
                
                self.log(f"INFO: Uploading metadata to: {s3_metadata_path}")
                
                result = subprocess.run(
                    aws_cmd,
                    env=env,
                    capture_output=True,
                    text=True,
                    timeout=60
                )
                
                if result.returncode == 0:
                    self.log("SUCCESS: Metadata file uploaded to S3")
                    self.log(f"INFO: Metadata location: {s3_metadata_path}")
                    return True
                else:
                    self.log(f"ERROR: Failed to upload metadata to S3: {result.stderr}")
                    return False
                    
            finally:
                # Clean up temporary file
                try:
                    os.unlink(temp_file_path)
                except:
                    pass
                    
        except Exception as e:
            self.log(f"ERROR: Failed to store S3 metadata file: {e}")
            return False

    def scan_s3_for_images(self):
        """Scan S3 repository for existing image metadata and populate database"""
        try:
            self.log("INFO: Scanning S3 for existing image metadata...")
            
            # Only scan if S3 is configured
            s3_config = self.db.get_s3_config()
            if not s3_config:
                self.log("INFO: No S3 configuration - skipping image scan")
                return
                
            # Set up environment for AWS CLI
            env = os.environ.copy()
            env['AWS_ACCESS_KEY_ID'] = s3_config['s3_access_key']
            env['AWS_SECRET_ACCESS_KEY'] = s3_config['s3_secret_key']
            if s3_config.get('s3_endpoint') and not s3_config['s3_endpoint'].startswith('s3.'):
                env['AWS_ENDPOINT_URL'] = f"https://{s3_config['s3_endpoint']}"
            
            # List all metadata files in the bucket
            bucket = s3_config['s3_bucket']
            aws_cmd = [
                "aws", "s3", "ls", f"s3://{bucket}/", 
                "--recursive", "--output", "json"
            ]
            
            result = subprocess.run(
                aws_cmd,
                env=env,
                capture_output=True,
                text=True,
                timeout=120
            )
            
            if result.returncode != 0:
                self.log(f"WARNING: Failed to list S3 objects: {result.stderr}")
                return
                
            # Find metadata files
            metadata_files = []
            for line in result.stdout.strip().split('\n'):
                if not line.strip():
                    continue
                try:
                    obj = json.loads(line)
                    key = obj.get('Key', '')
                    if '/metadata/' in key and key.endswith('.json'):
                        metadata_files.append(key)
                except:
                    continue
            
            self.log(f"INFO: Found {len(metadata_files)} metadata files")
            
            discovered_clients = set()
            discovered_sites = set()
            
            # Process each metadata file
            for metadata_key in metadata_files:
                try:
                    # Download metadata file
                    download_cmd = [
                        "aws", "s3", "cp", f"s3://{bucket}/{metadata_key}", "-"
                    ]
                    
                    download_result = subprocess.run(
                        download_cmd,
                        env=env,
                        capture_output=True,
                        text=True,
                        timeout=30
                    )
                    
                    if download_result.returncode == 0:
                        metadata = json.loads(download_result.stdout)
                        
                        # Extract client and site information
                        client_info = metadata.get('client_info', {})
                        client_uuid = client_info.get('client_uuid')
                        client_name = client_info.get('client_name')
                        site_uuid = client_info.get('site_uuid')
                        site_name = client_info.get('site_name')
                        role = client_info.get('role', 'OP')
                        
                        # Add client to database if not exists
                        if client_uuid and client_name:
                            try:
                                # Check if client exists
                                existing_clients = self.db.get_clients()
                                client_exists = any(cid == client_uuid for cid, _, _, _ in existing_clients)
                                
                                if not client_exists:
                                    # Create client with discovered info
                                    client_short = client_name.lower().replace(' ', '-')[:20]
                                    description = f"Auto-discovered from S3 image metadata"
                                    
                                    # Use add_client but with specific UUID
                                    cursor = self.db.connection.cursor()
                                    cursor.execute(
                                        "INSERT INTO clients (id, name, short_name, description) VALUES (?, ?, ?, ?)",
                                        (client_uuid, client_name, client_short, description)
                                    )
                                    self.db.connection.commit()
                                    
                                    discovered_clients.add(client_name)
                                    self.log(f"INFO: Discovered client: {client_name}")
                                    
                                # Add site to database if not exists
                                if site_uuid and site_name:
                                    existing_sites = self.db.get_sites(client_uuid)
                                    site_exists = any(sid == site_uuid for sid, _, _, _, _, _ in existing_sites)
                                    
                                    if not site_exists:
                                        site_short = site_name.lower().replace(' ', '-')[:20]
                                        description = f"Auto-discovered from S3 image metadata"
                                        
                                        cursor.execute(
                                            "INSERT INTO sites (id, client_id, name, short_name, description) VALUES (?, ?, ?, ?, ?)",
                                            (site_uuid, client_uuid, site_name, site_short, description)
                                        )
                                        self.db.connection.commit()
                                        
                                        discovered_sites.add(f"{client_name}/{site_name}")
                                        self.log(f"INFO: Discovered site: {client_name}/{site_name}")
                                        
                            except Exception as e:
                                self.log(f"WARNING: Failed to process client/site from {metadata_key}: {e}")
                        
                        # Store image metadata in database
                        backup_uuid = metadata.get('backup_uuid')
                        if backup_uuid:
                            try:
                                # Check if image already exists
                                existing_images = self.db.get_images()
                                image_exists = any(img_id == backup_uuid for img_id, _, _, _, _, _, _, _, _ in existing_images)
                                
                                if not image_exists:
                                    # Create image record
                                    self.db.create_image(
                                        image_id=backup_uuid,
                                        client_id=client_uuid or "unknown",
                                        site_id=site_uuid or "unknown", 
                                        role=role,
                                        repository_path=f"s3://{bucket}/{metadata_key}",
                                        repository_size_gb=0
                                    )
                                    
                                    # Store full metadata as JSON
                                    self.db.save_image_metadata(backup_uuid, json.dumps(metadata))
                                    
                            except Exception as e:
                                self.log(f"WARNING: Failed to store image metadata for {backup_uuid}: {e}")
                                
                except Exception as e:
                    self.log(f"WARNING: Failed to process metadata file {metadata_key}: {e}")
                    continue
            
            # Update UI if we discovered new clients/sites
            if discovered_clients or discovered_sites:
                self.log(f"INFO: Image scan complete - discovered {len(discovered_clients)} clients, {len(discovered_sites)} sites")
                
                # Refresh client/site data in the UI
                if hasattr(self, 'refresh_client_site_data'):
                    self.root.after(0, self.refresh_client_site_data)
            else:
                self.log("INFO: Image scan complete - no new clients or sites discovered")
                
        except Exception as e:
            self.log(f"ERROR: Failed to scan S3 for images: {e}")

    def start_vss_restic_creation_thread(self):
        """Starts the VSS + Restic backup process in a new thread."""
        # Disable buttons
        if hasattr(self, 'vss_create_button'):
            self.vss_create_button.config(state="disabled")
        if hasattr(self, 'direct_create_button'):
            self.direct_create_button.config(state="disabled")
        
        thread = threading.Thread(target=self.vss_restic_creation_worker)
        thread.daemon = True
        thread.start()

    def vss_restic_creation_worker(self):
        """Worker function for VSS + Restic backup creation."""
        try:
            success = self.create_vss_restic_backup()
            if success:
                self.log("=== VSS + RESTIC BACKUP COMPLETED SUCCESSFULLY ===")
            else:
                self.log("=== VSS + RESTIC BACKUP FAILED ===")
        except Exception as e:
            self.log(f"FATAL: VSS + Restic worker failed: {e}")
        finally:
            # Re-enable buttons
            if hasattr(self, 'vss_create_button'):
                self.vss_create_button.config(state="normal")
            if hasattr(self, 'direct_create_button'):
                self.direct_create_button.config(state="normal")

    def capture_with_dism(self, source_path, destination_path, method_name, use_enhanced_settings=False):
        """
        Helper method to capture WIM using DISM with different approaches
        Enhanced to handle shadow copy lifetime issues
        """
        try:
            self.log(f"INFO: DISM capture via {method_name}")
            self.log(f"INFO: Source: {source_path}")
            self.log(f"INFO: Destination: {destination_path}")
            
            # Build DISM command with appropriate settings
            dism_cmd = [
                "dism", "/capture-image",
                f"/imagefile:{destination_path}",
                f"/capturedir:{source_path}",
                "/name:System Image (VSS)",
                "/description:Windows system image captured via VSS + DISM"
            ]
            
            if use_enhanced_settings:
                # Enhanced settings for direct shadow copy access
                dism_cmd.extend(["/compress:fast", "/verify"])
                # Remove /bootable as it can cause issues with shadow copies
                self.log("INFO: Using enhanced DISM settings for shadow copy")
            else:
                # Standard settings
                dism_cmd.extend(["/compress:fast", "/verify"])
            
            self.log(f"COMMAND: {' '.join(dism_cmd)}")
            self.log(f"INFO: Starting DISM capture via {method_name}...")
            
            # Execute DISM with better process handling
            dism_proc = subprocess.Popen(
                dism_cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,  # Separate stderr
                text=True,
                encoding='utf-8',
                errors='ignore'
            )
            
            # Stream output with better error handling
            output_lines = []
            error_lines = []
            
            try:
                # Read output in a non-blocking way to prevent hanging
                import select
                import sys
                
                if sys.platform == 'win32':
                    # Windows doesn't support select on pipes, use threading
                    import threading
                    import queue
                    
                    def read_output(pipe, output_queue, line_list):
                        try:
                            for line in iter(pipe.readline, ''):
                                if not line:
                                    break
                                line = line.strip()
                                if line:
                                    output_queue.put(('stdout', line))
                                    line_list.append(line)
                        except Exception as e:
                            output_queue.put(('error', str(e)))
                    
                    def read_errors(pipe, output_queue, line_list):
                        try:
                            for line in iter(pipe.readline, ''):
                                if not line:
                                    break
                                line = line.strip()
                                if line:
                                    output_queue.put(('stderr', line))
                                    line_list.append(line)
                        except Exception as e:
                            output_queue.put(('error', str(e)))
                    
                    output_queue = queue.Queue()
                    
                    # Start threads to read stdout and stderr
                    stdout_thread = threading.Thread(target=read_output, args=(dism_proc.stdout, output_queue, output_lines))
                    stderr_thread = threading.Thread(target=read_errors, args=(dism_proc.stderr, output_queue, error_lines))
                    
                    stdout_thread.daemon = True
                    stderr_thread.daemon = True
                    
                    stdout_thread.start()
                    stderr_thread.start()
                    
                    # Process output as it comes in
                    last_progress = 0
                    stalled_count = 0
                    
                    while dism_proc.poll() is None:
                        try:
                            # Check for new output with timeout
                            stream_type, line = output_queue.get(timeout=30)  # 30 second timeout
                            
                            if stream_type == 'stdout':
                                self.log(line)
                                
                                # Check for progress indicators
                                if '%' in line and '[' in line:
                                    try:
                                        # Extract percentage
                                        percent_str = line.split('%')[0].split()[-1]
                                        current_progress = float(percent_str)
                                        if current_progress > last_progress:
                                            last_progress = current_progress
                                            stalled_count = 0
                                        else:
                                            stalled_count += 1
                                    except:
                                        pass
                                
                                # Check for critical errors that indicate shadow copy issues
                                if any(error in line.lower() for error in [
                                    "the handle is invalid",
                                    "access is denied", 
                                    "the system cannot find the path specified",
                                    "the filename, directory name, or volume label syntax is incorrect"
                                ]):
                                    self.log(f"ERROR: Critical error detected during DISM {method_name}: {line}")
                                    self.log(f"ERROR: Shadow copy became invalid during capture")
                                    self.log(f"ERROR: This is a common VSS issue - shadow copy lifetime expired")
                                    self.log(f"SOLUTION: Try these steps:")
                                    self.log(f"  1. Close all applications and background processes")
                                    self.log(f"  2. Temporarily disable antivirus real-time protection")
                                    self.log(f"  3. Increase virtual memory/page file size")
                                    self.log(f"  4. Stop Windows Search service temporarily")
                                    self.log(f"  5. Run: net stop themes (reduces memory usage)")
                                    self.log(f"  6. Try smaller capture chunks if possible")
                                    dism_proc.terminate()
                                    return False
                                    
                            elif stream_type == 'stderr':
                                self.log(f"STDERR: {line}")
                                
                            elif stream_type == 'error':
                                self.log(f"THREAD ERROR: {line}")
                                
                        except queue.Empty:
                            # Timeout - check if process is still alive
                            if dism_proc.poll() is None:
                                self.log("WARNING: DISM process appears stalled, checking status...")
                                stalled_count += 1
                                if stalled_count > 10:  # 5 minutes of stalling
                                    self.log("ERROR: DISM process has been stalled for too long, terminating")
                                    dism_proc.terminate()
                                    return False
                            else:
                                break
                    
                    # Wait for threads to finish
                    stdout_thread.join(timeout=5)
                    stderr_thread.join(timeout=5)
                    
                else:
                    # Non-Windows systems (fallback)
                    if dism_proc.stdout:
                        for line in iter(dism_proc.stdout.readline, ''):
                            if not line:
                                break
                            line = line.strip()
                            if line:
                                self.log(line)
                                output_lines.append(line)
                
                # Wait for process to complete
                dism_proc.wait()
                
            except Exception as e:
                self.log(f"ERROR: Exception during DISM {method_name}: {e}")
                try:
                    dism_proc.terminate()
                except:
                    pass
                return False
            
            # Check return code
            if dism_proc.returncode == 0:
                self.log(f"SUCCESS: DISM {method_name} completed successfully!")
                self.log(f"SUCCESS: WIM image saved to: {destination_path}")
                
                # Verify file was created and get size
                dest_file = Path(destination_path)
                if dest_file.exists():
                    size_gb = dest_file.stat().st_size / (1024**3)
                    self.log(f"SUCCESS: Created WIM file ({size_gb:.1f} GB)")
                    return True
                else:
                    self.log(f"ERROR: DISM reported success but file not found: {destination_path}")
                    return False
            else:
                self.log(f"ERROR: DISM {method_name} failed with exit code: {dism_proc.returncode}")
                if error_lines:
                    self.log(f"ERROR: DISM errors: {'; '.join(error_lines)}")
                return False
                
        except Exception as e:
            self.log(f"ERROR: DISM {method_name} failed with exception: {e}")
            return False

    def create_vhdx_dialog(self):
        """Create VHDX with repository restore and VM setup"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Create VHDX with Repository Restore")
        dialog.geometry("900x700")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # Title
        ttk.Label(dialog, text="Create VHDX with Repository Restore", 
                 font=("TkDefaultFont", 14, "bold")).pack(pady=20)
        ttk.Label(dialog, text="Create a new VHDX, partition it, restore from restic repository, and create VM", 
                 font=("TkDefaultFont", 10)).pack(pady=5)
        
        # Repository selection
        repo_frame = ttk.LabelFrame(dialog, text="Repository Selection", padding="10")
        repo_frame.pack(fill="x", padx=20, pady=10)
        
        ttk.Label(repo_frame, text="Select repository to restore from:").pack(anchor="w")
        
        repo_var = tk.StringVar()
        repo_combo = ttk.Combobox(repo_frame, textvariable=repo_var, width=70, state="readonly")
        repo_combo.pack(fill="x", pady=5)
        
        # Load repositories
        repositories = []
        try:
            images = self.db.get_images()
            for image in images:
                if len(image) >= 10:  # Ensure we have all fields
                    client_name = self.db.get_client_name(image[1]) or "Unknown"
                    site_name = self.db.get_site_name(image[2]) or "Unknown"
                    role = image[3] or "unknown"
                    repo_path = image[4] or ""
                    display_name = f"{client_name}/{site_name}/{role} - {repo_path}"
                    repositories.append((display_name, image))
        except Exception as e:
            self.log_step2(f"Error loading repositories: {e}")
        
        repo_combo['values'] = [repo[0] for repo in repositories]
        
        # VHDX Configuration
        vhdx_frame = ttk.LabelFrame(dialog, text="VHDX Configuration", padding="10")
        vhdx_frame.pack(fill="x", padx=20, pady=10)
        
        # Size configuration
        size_container = ttk.Frame(vhdx_frame)
        size_container.pack(fill="x", pady=5)
        
        ttk.Label(size_container, text="VHDX Size (GB):").pack(side="left")
        size_var = tk.IntVar(value=256)
        size_spinbox = ttk.Spinbox(size_container, from_=64, to=2048, textvariable=size_var, width=10)
        size_spinbox.pack(side="left", padx=(10, 0))
        
        ttk.Label(size_container, text="(Default: 256 GB, Dynamic allocation)").pack(side="left", padx=(10, 0))
        
        # VHDX name
        name_container = ttk.Frame(vhdx_frame)
        name_container.pack(fill="x", pady=5)
        
        ttk.Label(name_container, text="VHDX Name:").pack(side="left")
        name_var = tk.StringVar(value="RestoreImage")
        name_entry = ttk.Entry(name_container, textvariable=name_var, width=30)
        name_entry.pack(side="left", padx=(10, 0))
        
        # VM Configuration
        vm_frame = ttk.LabelFrame(dialog, text="Virtual Machine Configuration", padding="10")
        vm_frame.pack(fill="x", padx=20, pady=10)
        
        vm_info = ttk.Label(vm_frame, text="VM will be created with: 4 GB RAM, 4 CPUs, Secure Boot ON, TPM ON")
        vm_info.pack(anchor="w")
        
        vm_name_container = ttk.Frame(vm_frame)
        vm_name_container.pack(fill="x", pady=5)
        
        ttk.Label(vm_name_container, text="VM Name:").pack(side="left")
        vm_name_var = tk.StringVar(value="RestoreVM")
        vm_name_entry = ttk.Entry(vm_name_container, textvariable=vm_name_var, width=30)
        vm_name_entry.pack(side="left", padx=(10, 0))
        
        # Process steps preview
        steps_frame = ttk.LabelFrame(dialog, text="Process Steps", padding="10")
        steps_frame.pack(fill="x", padx=20, pady=10)
        
        steps_text = """1. Create dynamic VHDX file
2. Mount VHDX and partition (4GB EFI, 4GB Recovery, Rest for OS)
3. Format partitions (EFI: FAT32, Recovery: NTFS, OS: NTFS)
4. Restore selected repository to OS partition
5. Unmount VHDX
6. Create Hyper-V VM with Secure Boot and TPM
7. Attach VHDX as boot drive"""
        
        ttk.Label(steps_frame, text=steps_text, justify="left").pack(anchor="w")
        
        # Progress area
        progress_frame = ttk.LabelFrame(dialog, text="Progress", padding="10")
        progress_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        status_var = tk.StringVar(value="Ready to create VHDX...")
        ttk.Label(progress_frame, textvariable=status_var).pack(anchor="w")
        
        progress_bar = ttk.Progressbar(progress_frame, mode='indeterminate')
        progress_bar.pack(fill="x", pady=5)
        
        # Progress log
        log_text = scrolledtext.ScrolledText(
            progress_frame, 
            height=8, 
            font=("Consolas", 9),
            bg="#1e1e1e", 
            fg="#ffffff", 
            insertbackground="#ffffff"
        )
        log_text.pack(fill="both", expand=True, pady=5)
        
        # Buttons
        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill="x", padx=20, pady=20)
        
        def start_vhdx_creation():
            """Start the VHDX creation process"""
            repo_selection = repo_var.get()
            vhdx_size = size_var.get()
            vhdx_name = name_var.get().strip()
            vm_name = vm_name_var.get().strip()
            
            # Validation
            if not repo_selection:
                messagebox.showerror("Error", "Please select a repository to restore from")
                return
            
            if not vhdx_name:
                messagebox.showerror("Error", "Please enter a VHDX name")
                return
            
            if not vm_name:
                messagebox.showerror("Error", "Please enter a VM name")
                return
            
            # Find selected repository data
            selected_repo = None
            for display_name, repo_data in repositories:
                if display_name == repo_selection:
                    selected_repo = repo_data
                    break
            
            if not selected_repo:
                messagebox.showerror("Error", "Invalid repository selection")
                return
            
            # Confirm creation
            if not messagebox.askyesno("Confirm VHDX Creation", 
                                     f"Create VHDX and VM with the following settings?\n\n"
                                     f"Repository: {repo_selection}\n"
                                     f"VHDX Size: {vhdx_size} GB\n"
                                     f"VHDX Name: {vhdx_name}\n"
                                     f"VM Name: {vm_name}\n\n"
                                     "This process may take 30+ minutes.\n\n"
                                     "Continue?"):
                return
            
            # Start creation process
            progress_bar.start()
            status_var.set("Starting VHDX creation...")
            
            def log_to_dialog(message):
                """Log message to the dialog"""
                timestamp = datetime.now().strftime("%H:%M:%S")
                formatted_message = f"[{timestamp}] {message}\n"
                log_text.insert(tk.END, formatted_message)
                log_text.see(tk.END)
                log_text.update()
                self.log_step2(message)
            
            def creation_thread():
                try:
                    success = self.perform_vhdx_creation_workflow(
                        selected_repo, vhdx_size, vhdx_name, vm_name, 
                        log_to_dialog, status_var
                    )
                    
                    dialog.after(0, lambda: creation_complete(success))
                    
                except Exception as e:
                    dialog.after(0, lambda: creation_failed(str(e)))
            
            def creation_complete(success):
                progress_bar.stop()
                if success:
                    status_var.set("VHDX and VM created successfully!")
                    log_to_dialog("✓ VHDX creation and VM setup completed successfully!")
                    messagebox.showinfo("Success", "VHDX and VM created successfully!")
                else:
                    status_var.set("VHDX creation failed!")
                    log_to_dialog("✗ VHDX creation failed!")
            
            def creation_failed(error):
                progress_bar.stop()
                status_var.set("VHDX creation failed!")
                log_to_dialog(f"✗ FATAL ERROR: {error}")
                messagebox.showerror("Error", f"VHDX creation failed:\n{error}")
            
            threading.Thread(target=creation_thread, daemon=True).start()
        
        start_btn = ttk.Button(button_frame, text="🚀 Create VHDX & VM", command=start_vhdx_creation)
        start_btn.pack(side="left", padx=5)
        
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side="left", padx=5)

    def perform_vhdx_creation_workflow(self, repo_data, vhdx_size_gb, vhdx_name, vm_name, log_func, status_var):
        """Perform the complete VHDX creation workflow"""
        try:
            log_func("Starting VHDX creation workflow...")
            
            # Get working directory
            working_dir = self.db.get_working_vhdx_directory()
            vhdx_path = working_dir / f"{vhdx_name}.vhdx"
            
            log_func(f"Working directory: {working_dir}")
            log_func(f"VHDX path: {vhdx_path}")
            log_func(f"Repository: {repo_data[4]}")  # repository_path
            
            # Step 1: Create VHDX
            status_var.set("Creating VHDX file...")
            log_func("Step 1: Creating VHDX file...")
            if not self.create_dynamic_vhdx(vhdx_path, vhdx_size_gb, log_func):
                raise Exception("Failed to create VHDX file")
            
            # Step 2: Mount and partition VHDX
            status_var.set("Mounting and partitioning VHDX...")
            log_func("Step 2: Mounting and partitioning VHDX...")
            drive_letter = self.mount_and_partition_vhdx(vhdx_path, log_func)
            if not drive_letter:
                raise Exception("Failed to mount and partition VHDX")
            
            # Step 3: Restore from repository
            status_var.set("Restoring from repository...")
            log_func("Step 3: Restoring from repository...")
            if not self.restore_repository_to_vhdx_partition(repo_data, drive_letter, log_func):
                raise Exception("Failed to restore repository to VHDX")
            
            # Step 4: Unmount VHDX
            status_var.set("Unmounting VHDX...")
            log_func("Step 4: Unmounting VHDX...")
            if not self.unmount_vhdx(vhdx_path, log_func):
                log_func("Warning: Failed to cleanly unmount VHDX, but continuing...")
            
            # Step 5: Create VM
            status_var.set("Creating VM...")
            log_func("Step 5: Creating Hyper-V VM...")
            if not self.create_hyperv_vm(vm_name, vhdx_path, log_func):
                raise Exception("Failed to create Hyper-V VM")
            
            log_func("✓ VHDX creation workflow completed successfully!")
            status_var.set("Completed successfully!")
            return True
            
        except Exception as e:
            log_func(f"✗ VHDX creation workflow failed: {str(e)}")
            status_var.set("Failed!")
            return False

    def create_dynamic_vhdx(self, vhdx_path, size_gb, log_func):
        """Create a dynamic VHDX file"""
        try:
            log_func(f"Creating dynamic VHDX: {vhdx_path}")
            log_func(f"Size: {size_gb} GB")
            
            # Ensure parent directory exists
            vhdx_path.parent.mkdir(parents=True, exist_ok=True)
            
            # Create diskpart script for VHDX creation
            script_content = f"""create vdisk file="{vhdx_path}" maximum={size_gb * 1024} type=expandable
exit
"""
            
            script_path = vhdx_path.parent / "create_vhdx.txt"
            with open(script_path, 'w') as f:
                f.write(script_content)
            
            log_func(f"Created diskpart script: {script_path}")
            
            # Run diskpart
            cmd = ["diskpart", "/s", str(script_path)]
            log_func(f"Running: {' '.join(cmd)}")
            
            result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
            
            log_func(f"Diskpart return code: {result.returncode}")
            if result.stdout:
                log_func(f"STDOUT: {result.stdout}")
            if result.stderr:
                log_func(f"STDERR: {result.stderr}")
            
            # Clean up script
            try:
                script_path.unlink()
            except:
                pass
            
            if result.returncode == 0 and vhdx_path.exists():
                log_func(f"✓ VHDX created successfully: {vhdx_path}")
                return True
            else:
                log_func(f"✗ Failed to create VHDX")
                return False
                
        except Exception as e:
            log_func(f"✗ Error creating VHDX: {e}")
            return False

    def mount_and_partition_vhdx(self, vhdx_path, log_func):
        """Mount VHDX and create partitions (EFI, Recovery, OS)"""
        try:
            log_func("Mounting and partitioning VHDX...")
            
            # Create diskpart script for mounting and partitioning
            script_content = f"""select vdisk file="{vhdx_path}"
attach vdisk
create partition efi size=4096
active
format fs=fat32 quick label="EFI"
assign letter=E
create partition msr size=128
create partition primary size=4096
format fs=ntfs quick label="Recovery"
assign letter=R
create partition primary
format fs=ntfs quick label="OS"
assign letter=O
list partition
exit
"""
            
            script_path = vhdx_path.parent / "partition_vhdx.txt"
            with open(script_path, 'w') as f:
                f.write(script_content)
            
            log_func(f"Created partition script: {script_path}")
            
            # Run diskpart
            cmd = ["diskpart", "/s", str(script_path)]
            log_func(f"Running: {' '.join(cmd)}")
            
            result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
            
            log_func(f"Diskpart return code: {result.returncode}")
            if result.stdout:
                log_func(f"STDOUT: {result.stdout}")
            if result.stderr:
                log_func(f"STDERR: {result.stderr}")
            
            # Clean up script
            try:
                script_path.unlink()
            except:
                pass
            
            if result.returncode == 0:
                log_func("✓ VHDX mounted and partitioned successfully")
                log_func("✓ EFI partition: E: (4GB, FAT32)")
                log_func("✓ Recovery partition: R: (4GB, NTFS)")
                log_func("✓ OS partition: O: (Remaining space, NTFS)")
                return "O"  # Return OS drive letter
            else:
                log_func(f"✗ Failed to mount and partition VHDX")
                return None
                
        except Exception as e:
            log_func(f"✗ Error mounting and partitioning VHDX: {e}")
            return None

    def restore_repository_to_vhdx_partition(self, repo_data, drive_letter, log_func):
        """Restore repository to the VHDX OS partition"""
        try:
            log_func(f"Restoring repository to {drive_letter}: drive...")
            
            # Get repository details
            repo_path = repo_data[4]  # repository_path
            repo_password = repo_data[8] if len(repo_data) > 8 else None  # restic_password
            
            if not repo_password:
                log_func("✗ No repository password found")
                return False
            
            log_func(f"Repository path: {repo_path}")
            
            # Get restic binary
            restic_exe = self.download_restic()
            if not restic_exe:
                log_func("✗ Could not obtain restic binary")
                return False
            
            log_func(f"Using restic: {restic_exe}")
            
            # Set environment
            os.environ['RESTIC_REPOSITORY'] = str(repo_path)
            os.environ['RESTIC_PASSWORD'] = repo_password
            
            # Find latest snapshot
            log_func("Finding latest snapshot...")
            snapshots_cmd = [restic_exe, "snapshots", "--json"]
            result = subprocess.run(snapshots_cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
            
            if result.returncode != 0:
                log_func(f"✗ Failed to query snapshots: {result.stderr}")
                return False
            
            try:
                snapshots = json.loads(result.stdout) if result.stdout.strip() else []
                if not snapshots:
                    log_func("✗ No snapshots found in repository")
                    return False
                
                latest_snapshot = snapshots[-1]['short_id']
                log_func(f"✓ Using latest snapshot: {latest_snapshot}")
                
            except (json.JSONDecodeError, KeyError, IndexError) as e:
                log_func(f"✗ Error parsing snapshots: {e}")
                return False
            
            # Restore to OS partition
            target_path = f"{drive_letter}:\\"
            log_func(f"Restoring to: {target_path}")
            
            restore_cmd = [restic_exe, "restore", latest_snapshot, "--target", target_path]
            log_func(f"Running: {' '.join(restore_cmd)}")
            
            result = subprocess.run(restore_cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
            
            log_func(f"Restore return code: {result.returncode}")
            if result.stdout:
                log_func(f"STDOUT: {result.stdout}")
            if result.stderr:
                log_func(f"STDERR: {result.stderr}")
            
            if result.returncode == 0:
                log_func(f"✓ Repository restored successfully to {drive_letter}:")
                return True
            else:
                log_func(f"✗ Failed to restore repository")
                return False
                
        except Exception as e:
            log_func(f"✗ Error restoring repository: {e}")
            return False

    def unmount_vhdx(self, vhdx_path, log_func):
        """Unmount the VHDX"""
        try:
            log_func("Unmounting VHDX...")
            
            # Create diskpart script for unmounting
            script_content = f"""select vdisk file="{vhdx_path}"
detach vdisk
exit
"""
            
            script_path = vhdx_path.parent / "unmount_vhdx.txt"
            with open(script_path, 'w') as f:
                f.write(script_content)
            
            log_func(f"Created unmount script: {script_path}")
            
            # Run diskpart
            cmd = ["diskpart", "/s", str(script_path)]
            log_func(f"Running: {' '.join(cmd)}")
            
            result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
            
            log_func(f"Diskpart return code: {result.returncode}")
            if result.stdout:
                log_func(f"STDOUT: {result.stdout}")
            if result.stderr:
                log_func(f"STDERR: {result.stderr}")
            
            # Clean up script
            try:
                script_path.unlink()
            except:
                pass
            
            if result.returncode == 0:
                log_func("✓ VHDX unmounted successfully")
                return True
            else:
                log_func(f"✗ Failed to unmount VHDX")
                return False
                
        except Exception as e:
            log_func(f"✗ Error unmounting VHDX: {e}")
            return False

    def create_hyperv_vm(self, vm_name, vhdx_path, log_func):
        """Create Hyper-V VM with specified configuration"""
        try:
            log_func(f"Creating Hyper-V VM: {vm_name}")
            
            # Check if Hyper-V is available
            check_cmd = ["powershell", "-Command", "Get-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V"]
            result = subprocess.run(check_cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
            
            if "Enabled" not in result.stdout:
                log_func("⚠ Hyper-V may not be enabled. VM creation might fail.")
            
            # Create VM with PowerShell
            ps_script = f'''
# Create new VM
New-VM -Name "{vm_name}" -MemoryStartupBytes 4GB -Generation 2
            
# Set VM configuration
Set-VM -Name "{vm_name}" -ProcessorCount 4
Set-VM -Name "{vm_name}" -DynamicMemory -MemoryMinimumBytes 2GB -MemoryMaximumBytes 8GB
            
# Enable Secure Boot and TPM
Set-VMFirmware -VMName "{vm_name}" -EnableSecureBoot On
Set-VMSecurity -VMName "{vm_name}" -TpmEnabled $true
            
# Attach VHDX
Add-VMHardDiskDrive -VMName "{vm_name}" -Path "{vhdx_path}"
            
# Set boot order
$bootOrder = Get-VMFirmware -VMName "{vm_name}" | Select-Object -ExpandProperty BootOrder
$hardDrive = $bootOrder | Where-Object {{$_.Device -like "*Hard Drive*"}} | Select-Object -First 1
if ($hardDrive) {{
    Set-VMFirmware -VMName "{vm_name}" -FirstBootDevice $hardDrive.Device
}}
            
Write-Host "VM created successfully: {vm_name}"
'''
            
            cmd = ["powershell", "-Command", ps_script]
            log_func("Creating VM with PowerShell...")
            
            result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8', errors='ignore')
            
            log_func(f"PowerShell return code: {result.returncode}")
            if result.stdout:
                log_func(f"STDOUT: {result.stdout}")
            if result.stderr:
                log_func(f"STDERR: {result.stderr}")
            
            if result.returncode == 0:
                log_func(f"✓ Hyper-V VM created successfully: {vm_name}")
                log_func("✓ Configuration: 4GB RAM, 4 CPUs, Secure Boot ON, TPM ON")
                log_func(f"✓ VHDX attached: {vhdx_path}")
                return True
            else:
                log_func(f"✗ Failed to create Hyper-V VM")
                return False
                
        except Exception as e:
            log_func(f"✗ Error creating Hyper-V VM: {e}")
            return False

def main():
    # Check platform first
    if not check_platform():
        return
    
    # Check if running as administrator
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