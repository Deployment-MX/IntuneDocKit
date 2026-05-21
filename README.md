# IntuneDocKit

🚀 **Automated Microsoft Intune Documentation Generator**  
A PowerShell-based solution that automatically generates complete technical documentation for a Microsoft Intune tenant and exports it as a structured Word (`.docx`) document.

Based on the original project by matbe and modernized with:

- Automated App Registration
- Microsoft Graph integration
- Automatic permission assignment
- Admin Consent automation
- Modern Windows Forms GUI
- Improved export workflow and compatibility

---Last Update 5/21/2026 adding Microsoft Intune Suite workloads--------------------------

# Overview

Documenting Microsoft Intune environments manually is time-consuming, inconsistent, and difficult to maintain across projects.

**IntuneDocKit** automates the entire process by:

- Connecting securely to Microsoft Graph
- Extracting tenant configuration data
- Organizing the information into structured sections
- Generating a professional Word document ready for delivery or internal documentation

> ⚠️ All operations are **read-only**.  
> IntuneDocKit never modifies tenant configurations.

---

# Features

✅ Automated Intune documentation generation  
✅ Automated Azure App Registration creation  
✅ Automatic Microsoft Graph permission assignment  
✅ Automatic Admin Consent grant  
✅ Microsoft Graph API integration  
✅ GUI interface with real-time progress  
✅ Word document generation (`.docx`)  
✅ JSON export support  
✅ Modular architecture  
✅ Read-only tenant access  

---

# What IntuneDocKit Documents

IntuneDocKit exports more than **20 Intune areas**, including:

- Device Compliance Policies
- Device Configuration Profiles
- Enrollment Restrictions & Settings
- Windows Autopilot Profiles
- Mobile Applications
- App Protection Policies
- Conditional Access Policies
- RBAC Roles & Assignments
- PowerShell Scripts
- Windows Quality Update Policies
- Feature Update Policies
- Notification Templates
- Scope Tags
- Device Categories
- Apple VPP Tokens
- Apple Push Notification Certificate
- Managed Device Inventory Summary
- And more...

---

# Requirements

## Supported Environment

- Windows PowerShell 5.1
- Windows OS
- Internet connectivity

> ❌ PowerShell Core / PowerShell 7 is currently **not supported**

---

## Required Permissions

### Azure Permissions
An Azure account capable of:

- Creating App Registrations
- Granting Admin Consent

### Intune Permissions
An Intune account with:

- Read access to Intune configuration

---

# Automatic Module Installation

IntuneDocKit automatically installs all required modules during execution.

## Required Modules

- PSWriteWord
- PSWriteHTML
- ImportExcel
- Microsoft.Graph.Authentication
- Microsoft.Graph.Applications
- Microsoft.Graph.Identity.SignIns
- AzureAD `2.0.2.140`
- MSAL.PS

> ⚠️ Some environments may require running PowerShell as Administrator.

---

# Project Structure

```text
IntuneDocKit/
│
├── IntuneDocKit.ps1
│   GUI interface built with Windows Forms
│
├── MasterExport-MEM.ps1
│   Main orchestration engine
│
├── Export-MEMConfiguration.ps1
│   Export engine using Microsoft Graph API
│
├── Export-MEMConfiguration.xml
│   Main configuration file
│
├── MEMDocumentationTempl.docx
│   Word template used for exports
│
└── IntuneDocKit.exe
    Standalone executable version
```

---

# Execution Modes

## Mode 1 — Create New App Registration (First Run)

Used for new tenants.

### What it does

- Creates the App Registration
- Assigns permissions
- Grants Admin Consent
- Updates XML configuration
- Generates documentation

### Example

```powershell
.\MasterExport-MEM.ps1 `
-TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
-ClientName "Client Name"
```

---

## Mode 2 — Use Existing App Registration

Used after the App Registration already exists.

### What it does

- Skips App Registration creation
- Uses existing Client ID from XML
- Generates documentation directly

### Example

```powershell
.\MasterExport-MEM.ps1 `
-TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
-ClientName "Client Name" `
-SkipAppCreation
```

---

## Mode 3 — Create App Registration Only

Used when preparing tenant access without exporting documentation immediately.

### What it does

- Creates App Registration
- Assigns permissions
- Updates XML
- Does NOT generate documentation

### Example

```powershell
.\MasterExport-MEM.ps1 `
-TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
-OnlyCreateApp
```

---

# The Five-Step Process

## Step 1 — Module Validation

- Verifies required modules
- Installs missing modules automatically
- Configures NuGet provider
- Sets PSGallery as trusted

---

## Step 2 — App Registration

Creates an App Registration named:

```text
IntuneDocKit
```

Also:

- Enables Public Client Flows
- Creates Service Principal
- Reuses existing app if already present

> Only executed in Modes 1 and 3

---

## Step 3 — Permissions & Admin Consent

Automatically:

- Assigns Microsoft Graph delegated permissions
- Grants Admin Consent

Permissions are primarily read-only.

> Only executed in Modes 1 and 3

---

## Step 4 — XML Configuration Update

Updates:

- Tenant ID
- Client ID

inside:

```text
Export-MEMConfiguration.xml
```

---

## Step 5 — Export Engine

Launches a clean PowerShell session to avoid conflicts between:

- Microsoft.Graph
- MSAL.PS

Then:

- Authenticates user via Device Code Flow
- Executes Graph API calls
- Generates Word documentation

---

# Device Code Authentication

During authentication:

1. A one-time code is displayed
2. Open:

```text
https://microsoft.com/devicelogin
```

3. Enter the provided code
4. Authenticate with an Intune account
5. Export continues automatically

> The account used here only requires Intune read permissions.

---

# Microsoft Graph Permissions

The following delegated permissions are assigned automatically:

```text
openid
profile
offline_access
User.Read
DeviceManagementApps.Read.All
DeviceManagementConfiguration.Read.All
DeviceManagementManagedDevices.Read.All
DeviceManagementRBAC.Read.All
DeviceManagementServiceConfig.Read.All
DeviceManagementScripts.Read.All
Directory.Read.All
User.Read.All
Group.Read.All
Organization.Read.All
Application.Read.All
Policy.Read.All
AuditLog.Read.All
Agreement.Read.All
CloudPC.Read.All
Policy.ReadWrite.ConditionalAccess
```

> `Policy.ReadWrite.ConditionalAccess` is required to read Conditional Access policies through delegated Graph API access.

---

# XML Configuration

The file:

```text
Export-MEMConfiguration.xml
```

controls export behavior.

## Key Settings

| Setting | Description |
|---|---|
| Tenant | Tenant ID |
| ClientId | App Registration Client ID |
| Document | Enable/disable Word generation |
| DocumentName | Output file name |
| Export | Enable JSON exports |
| ExportPath | Output directory |
| graphApiVersion | Graph API version |

---

## Section Control

Each Intune section can be enabled or disabled individually:

```xml
<Process>
    <CompliancePolicies>True</CompliancePolicies>
    <ConditionalAccess>True</ConditionalAccess>
</Process>
```

---

# Output

Generated files are typically saved to:

```text
C:\temp
```

The generated document includes:

- Cover page
- Client information
- Export date
- Structured Intune sections
- Tables and configuration details

The document opens automatically when completed.

---

# GUI Interface

Running:

```text
IntuneDocKit.exe
```

or

```powershell
.\IntuneDocKit.ps1
```

launches the graphical interface.

## GUI Features

- Client Name input
- Tenant ID input
- User UPN input
- Execution Mode selector
- Real-time console output
- Five-step progress bar
- Copy PowerShell command button
- Device Code copy button

---

# Security Notes

## Read-Only Operations

IntuneDocKit:

✅ Reads tenant configuration  
✅ Generates documentation  
❌ Does NOT modify Intune settings  
❌ Does NOT create policies  
❌ Does NOT change assignments  

---

# Credits

This project is based on the original work by matbe.

Original repository:

https://github.com/matbe/MEMDocumentAndExporter

The project was modernized and extended with:

- Automated App Registration
- Admin Consent automation
- Microsoft Graph integration
- GUI enhancements
- Improved export workflow

---

# Author

**Jesús Octavio Rodríguez**  
Microsoft MVP • Windows Insider MVP • Microsoft Certified Trainer

Specialized in:

- Microsoft Intune
- Windows 365
- Azure Virtual Desktop
- SCCM/MECM
- Enterprise Mobility
- Modern Endpoint Management
