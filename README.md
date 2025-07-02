# Windows Image Preparation Tool

A professional-grade Windows system imaging and backup tool using modern Restic backup engine with S3 cloud storage integration.

## 🎯 Features

- **Development Mode**: Create development images with S3 cloud storage
- **Production Mode**: Create production-ready deployment images
- **Volume Shadow Copy (VSS)**: Consistent backups of live Windows systems
- **S3 Integration**: Store images securely in cloud storage (Backblaze B2, AWS S3, etc.)
- **Client/Site Organization**: Organize images by client and site
- **Secure Password Management**: Automatic generation and storage of repository passwords
- **Modern UI**: Easy-to-use graphical interface

## 📋 Requirements

- **Windows 10/11** (Administrator privileges required)
- **Python 3.8+** (64-bit recommended)
- **S3-compatible storage** (Backblaze B2, AWS S3, MinIO, etc.)

## 🚀 Quick Setup Guide

### Step 1: Install Python 3 (64-bit)

The easiest way to install Python is using **Ninite**:

1. Go to [ninite.com](https://ninite.com)
2. Check the box for **"Python 3.x"** 
3. Click **"Get Your Ninite"** to download the installer
4. Run the downloaded `Ninite Python 3.x Installer.exe` as Administrator
5. Wait for installation to complete

**Alternative**: Download directly from [python.org](https://python.org) and make sure to check "Add Python to PATH" during installation.

### Step 2: Download the Image Creator Tool

1. Go to the GitHub repository: https://github.com/Gumbees/image-creator
2. Click the green **"Code"** button
3. Select **"Download ZIP"**
4. Extract the ZIP file to a folder (e.g., `C:\ImageCreator\`)

**Alternative (using Git):**
```cmd
git clone https://github.com/Gumbees/image-creator.git
cd image-creator
```

### Step 3: Install Dependencies

1. Open **Windows PowerShell** or **Command Prompt** as Administrator
2. Navigate to the extracted folder:
   ```cmd
   cd C:\ImageCreator\image-creator
   ```
3. Install required Python packages:
   ```cmd
   pip install -r requirements.txt
   ```

### Step 4: Run the Application

1. In the same PowerShell/Command Prompt window, run:
   ```cmd
   python windows_image_prep_gui.py
   ```
2. The Windows Image Preparation Tool will open

## 🔧 Using Development Mode

Development Mode is perfect for creating and managing development system images with S3 cloud storage.

### Initial Setup

1. **Launch the Tool**: Run `python windows_image_prep_gui.py` as Administrator
2. **Select Development Mode**: Click the **"🔧 DEVELOP CAPTURE"** button
3. **Configure S3 Storage**: Click the **"Configure S3..."** button

### S3 Configuration

You'll need S3-compatible storage credentials. Popular options:

**Backblaze B2** (Recommended - Free tier available):
- Endpoint: `s3.us-west-002.backblazeb2.com` (or your region)
- Bucket: Your bucket name
- Access Key: Your B2 Application Key ID
- Secret Key: Your B2 Application Key

**AWS S3**:
- Endpoint: `s3.amazonaws.com` (or regional endpoint)
- Bucket: Your S3 bucket name
- Access Key: Your AWS Access Key ID
- Secret Key: Your AWS Secret Access Key

### Setting Up Clients and Sites

1. **Add a Client**:
   - Click **"New Client"** button
   - Enter client details (Name, Short Name, Description)
   - Click **"Save"**

2. **Add a Site**:
   - Select your client from the dropdown
   - Click **"New Site"** button  
   - Enter site details (Name, Short Name, Description)
   - Click **"Save"**

3. **Create an Image Entry**:
   - Select your client and site
   - Choose a role (e.g., "workstation", "server", "domain-controller")
   - Click **"Create Image"**

### Creating Your First Backup

1. **Select Configuration**:
   - Client: Choose your client
   - Site: Choose your site
   - Image: Select the image entry you created
   - Role: Verify the role is correct

2. **Start Backup**:
   - Click the **"🚀 Start Development Backup"** button (bottom left)
   - If this is a new repository, you'll see a password dialog - **SAVE THIS PASSWORD!**
   - If repository exists, enter the existing password when prompted

3. **Monitor Progress**:
   - Watch the log area for backup progress
   - VSS (Volume Shadow Copy) will be used for consistent backups
   - First backup may take longer as it uploads all data

## 🔐 Important: Repository Passwords

- **Save your repository password immediately** when shown
- **Store it in a password manager** or secure location
- **Without the password, you cannot access your backups**
- Each client/site combination has a unique repository password

## 📁 S3 Storage Structure

Your S3 bucket will be organized as:
```
your-bucket/
├── metadata/                    # Central metadata storage
│   ├── backup-uuid-1.json       # Backup metadata
│   ├── backup-uuid-2.json
│   └── ...
└── client-uuid/                 # Client repository data
    └── [restic repository files]
```

## 🛠️ Troubleshooting

### "Administrator Required" Error
- Right-click PowerShell/Command Prompt and select **"Run as administrator"**
- The tool needs admin privileges for VSS (Volume Shadow Copy)

### Python Not Found
- Reinstall Python using Ninite and ensure it's added to PATH
- Or download from python.org with "Add Python to PATH" checked

### S3 Connection Issues
- Verify your S3 credentials are correct
- Check your internet connection
- Ensure your S3 bucket exists and is accessible

### "Repository Password Required"
- Enter the password that was shown when the repository was first created
- Check your password manager for the repository password
- Repository passwords are unique per client/site combination

## 🎯 Next Steps

After creating development images, you can:

1. **Browse Images**: View and manage your backup snapshots
2. **Restore Images**: Restore to VHDX files for VM deployment
3. **Production Mode**: Create production-ready deployment images
4. **Generalization**: Prepare images for deployment with sysprep

## 📖 Advanced Usage

For advanced features like production imaging, generalization, and deployment, see the additional documentation in the application's help system.

## 🤝 Support

- **GitHub Issues**: https://github.com/Gumbees/image-creator/issues
- **Documentation**: Check the in-app help and log messages
- **Community**: GitHub Discussions

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.