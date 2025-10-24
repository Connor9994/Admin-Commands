# Admin-Commands

![GitHub stars](https://img.shields.io/github/stars/Connor9994/Admin-Commands?style=social) ![GitHub forks](https://img.shields.io/github/forks/Connor9994/Admin-Commands?style=social) ![GitHub issues](https://img.shields.io/github/issues/Connor9994/Admin-Commands) 

![1](https://github.com/Connor9994/Admin-Commands/blob/main/Photos/1.png)

## Overview

This PowerShell script creates a graphical user interface (GUI) that serves as a quick-access hub for common Office 365 applications, Excel automation tasks, and administrative functions specific to TANYR Healthcare.

## Features

### 🏠 Main Page
- **SharePoint**: Quick access to TANYR Healthcare's SharePoint portal
- **Outlook**: Direct link to Outlook Web App
- Multiple placeholder buttons for additional functionality

### 📊 Excel Commands
- **Excel Link-Creation**: Web scraping utility that extracts data from web pages
- **Excel Edit**: Direct Excel automation for cell manipulation
- **Browser Edit**: Internet Explorer automation for web interactions
- **Save As Excel**: Google search automation with CSV export capabilities
- **Test Email**: SMTP email testing functionality

### 🔧 Admin Commands
- **List Users**: Office 365 user management portal
- **List Groups**: Office 365 group management
- **Usage Reports**: Office 365 usage analytics
- **Mail Rules**: Exchange transport rules management
- **Azure Portal**: Azure administration dashboard
- **Exchange Admin Center**: Exchange Online management
- **Security Admin Center**: Security and compliance center
- **SharePoint Admin Center**: SharePoint administration

## Technical Details

### Requirements
- **PowerShell 3.0 or later** (with compatibility check built-in)
- **Windows Forms** (.NET Framework)
- **Internet Explorer** (for web automation features)
- **Microsoft Excel** (for Excel automation features)
- **Office 365 credentials** (for admin functionalities)

### Dependencies
- System.Windows.Forms assembly
- Custom icon file (`.\Files\Logo.ico`)
- Help image files (`.\Files\Help1.png`, `.\Files\Help2.png`)

## Usage

1. Run the script in PowerShell
2. Use the dropdown to navigate between pages:
   - **Main Page**: Primary application launchers
   - **Excel Commands**: Data processing and automation tools
   - **Admin Commands**: Administrative portals (requires appropriate permissions)

### Key Functionality

#### Web Automation
- Automated Internet Explorer instances for web scraping
- Google search integration with result parsing
- SharePoint document interaction

#### Excel Integration
- Direct COM object interaction with Excel
- Cell manipulation and data entry
- CSV export capabilities

#### Email System
- SMTP email testing with Office 365 credentials
- Secure authentication handling

## Security Notes

⚠️ **Important Security Considerations:**
- Requires Office 365 administrative privileges for admin features
- Stores and uses credentials for email functionality
- Uses Internet Explorer automation which may have security implications
- Contains hardcoded organizational URLs specific to TANYR Healthcare

## Legacy Notes

This script was developed as an internal automation tool and demonstrates:
- Early PowerShell GUI development techniques
- Internet Explorer automation (now deprecated)
- COM object integration with Office applications
- Custom form creation without modern frameworks

## Compatibility

- Designed for Windows environments
- Optimized for Office 365 ecosystem
- May require updates for modern PowerShell versions
- Internet Explorer dependencies may need migration to modern browsers

---

*Note: This is a legacy automation tool that may require significant updates for current environments and security standards.*
