# Windows Image Preparation GUI (PYC)

A comprehensive Python-based GUI tool for creating, managing, and deploying Windows system images using modern backup technologies like Restic with S3 cloud storage support.

## 🚀 Quick Start with Ninite

### Step 1: Install Python Using Ninite

[Ninite](https://ninite.com/) is the fastest and safest way to install Python on Windows:

1. **Visit Ninite**: Go to [https://ninite.com/](https://ninite.com/)
2. **Select Python**: Check the "Python 3.x" box in the "Dev Tools" section
3. **Download**: Click "Get Your Ninite" to download the installer
4. **Run**: Double-click the downloaded `.exe` file and let Ninite install Python automatically
   - No need to worry about PATH variables or installation options
   - Ninite installs the latest stable version with pip included
   - Automatically handles Windows-specific configuration

### Step 2: Verify Installation

Open Command Prompt or PowerShell and verify Python is installed:

```cmd
python --version
pip --version
```

You should see output showing Python 3.8+ and pip version information.

### Step 3: Download and Setup the Application

1. **Clone or download** this repository to your Windows machine
2. **Navigate** to the project directory:
   ```cmd
   cd path\to\image-creator
   ```

3. **Install dependencies**:
   ```cmd
   pip install -r requirements.txt
   ```

### Step 4: Run the Application

```cmd
python windows_image_prep_gui.py
```

## 📋 System Requirements

### Minimum Requirements
- **OS**: Windows 10/11 (x64)
- **Python**: 3.8 or higher
- **RAM**: 4GB minimum, 8GB recommended
- **Storage**: 10GB free space minimum
- **Permissions**: Administrator privileges required

### Recommended Setup
- **OS**: Windows 11 with latest updates
- **Python**: 3.11+ (installed via Ninite)
- **RAM**: 16GB or more
- **Storage**: 100GB+ free space for image operations
- **Network**: High-speed internet for S3 operations

## 🔧 Dependencies

The application has minimal dependencies thanks to Python's rich standard library:

### Required Dependencies
- **requests** (≥2.28.0) - For S3 API communications and web requests

### Built-in Modules Used
- **tkinter** - GUI framework (included with Python)
- **sqlite3** - Database operations (included with Python)
- **subprocess** - Running external tools like Restic (included with Python)
- **pathlib** - Modern file path handling (included with Python)

All other modules (`os`, `platform`, `json`, `uuid`, etc.) are part of Python's standard library.

## 🎯 Features

### Modern Backup Technology
- **Restic Integration**: Uses Restic for reliable, incremental backups with native VSS support
- **S3 Cloud Storage**: Full support for AWS S3 and S3-compatible storage
- **UUIDv7 Tagging**: Time-ordered unique identifiers for image tracking
- **Secure Passwords**: Cryptographically secure repository passwords with password manager integration

### Workflow Modes
- **Development Mode**: For creating and testing images (temporary database storage)
- **Production Mode**: For permanent image management and deployment

### Professional Features
- **Client/Site Organization**: Multi-tenant image management
- **Role-Based Imaging**: Support for workstation, server, domain controller images
- **Incremental Backups**: Update existing images without full re-capture
- **Metadata Management**: JSON-based metadata for import/export capabilities

## 🚀 Usage

### Initial Setup
1. **First Run**: The application will prompt for initial configuration
2. **Workflow Mode**: Choose Development or Production mode
3. **S3 Configuration**: Configure cloud storage (optional but recommended)

### Creating System Images
1. **Step 1 - Create System Backup**: 
   - Select client and site organization
   - Choose repository type (Local or S3)
   - Run Restic-based backup with VSS support

2. **Step 2 - Professional Image & VM Management**:
   - Browse and manage existing images
   - Create VMs from images
   - Export to various formats

3. **Step 3 - Generalize & Cleanup**:
   - Run Sysprep for image generalization
   - Cleanup temporary files and prepare for deployment

### S3 Cloud Storage Setup
1. Navigate to the S3 configuration section
2. Enter your S3 credentials:
   - **Endpoint**: Your S3 server URL (e.g., `s3.amazonaws.com`)
   - **Bucket**: Target bucket name
   - **Access Key**: S3 access key ID
   - **Secret Key**: S3 secret access key
   - **Region**: AWS region (e.g., `us-east-1`)

## 🛠 Troubleshooting

### Common Issues

#### "Python is not recognized as an internal or external command"
- **Solution**: Reinstall Python using Ninite, which automatically configures PATH variables
- **Alternative**: Manually add Python to your system PATH

#### "Permission denied" errors
- **Solution**: Run Command Prompt or PowerShell as Administrator
- **Note**: Image operations require elevated privileges

#### VSS (Volume Shadow Service) errors
- **Solution**: Ensure VSS service is running: `net start vss`
- **Alternative**: Restart Windows and try immediately after boot

#### S3 connection issues
- **Check**: Verify internet connectivity and S3 credentials
- **Firewall**: Ensure Windows Firewall allows the application
- **Endpoint**: Verify S3 endpoint URL format (include `https://` if needed)

### Performance Tips
- **SSD Storage**: Use SSD drives for better performance
- **Network**: Use wired connection for large S3 uploads
- **Exclusions**: Review backup exclusions to reduce image size
- **Antivirus**: Add exclusions for the application and repository folders

## 📁 Project Structure

```
image-creator/
├── windows_image_prep_gui.py    # Main application file
├── requirements.txt             # Python dependencies
├── README.md                   # This file
├── .gitignore                  # Git exclusions
├── .github/                    # GitHub workflows (optional)
└── build-arm64.spec           # PyInstaller build spec (optional)
```

## 🔒 Security Considerations

- **Database Storage**: Development mode uses temporary storage (`%WINDIR%\Temp\pyc.db`)
- **Password Generation**: Uses Python's `secrets` module for cryptographically secure passwords
- **S3 Credentials**: Stored locally in encrypted database
- **Repository Passwords**: Auto-generated and stored securely with password manager reminders

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🆘 Support

For issues and support:
1. Check the troubleshooting section above
2. Review existing GitHub Issues
3. Create a new issue with detailed information about your problem

## 🙏 Acknowledgments

- **Restic**: Modern backup tool with excellent VSS support
- **Ninite**: Simplified software installation for Windows
- **Python Community**: For the excellent standard library and tkinter GUI framework

---

**Note**: This tool is designed for Windows system administrators and IT professionals. Please ensure you understand image management and deployment concepts before using in production environments.