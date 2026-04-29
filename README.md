IntuneDocKit
A PowerShell tool that automatically generates complete technical documentation for a Microsoft Intune tenant and saves it as a Word (.docx) file. Based on the original project by @matbe, modernized and enhanced with automated App Registration, permissions, and Admin Consent via Microsoft Graph API.

Description
IntuneDocKit solves one of the most common problems in Microsoft Intune administration: technical documentation. Manually documenting a tenant means reviewing dozens of configurations, copying values, organizing sections, and keeping the document up to date — a slow, error-prone process that's hard to standardize across projects.
IntuneDocKit automates the entire workflow. It connects to the tenant using Microsoft Graph API through a dedicated App Registration, extracts the complete Intune configuration, and generates a structured Word document ready to deliver to a client or use internally as a technical backup.
All operations are read-only. The script never modifies any tenant configuration.

What It Documents
IntuneDocKit extracts and documents more than 20 Intune sections, including:
Device compliance policies, device configurations, enrollment settings, Windows Autopilot profiles, mobile apps and app protection policies, Conditional Access policies, RBAC roles and their assignments, device scripts, Windows Quality Update and Feature Update profiles, notification message templates, Scope Tags, device categories, Apple VPP tokens, Apple Push Notification certificate, and a general summary of the managed device inventory.

Requirements

PowerShell 5.1 on Windows. Not compatible with PowerShell Core or PowerShell 7.
An Azure account with permissions to create App Registrations and grant Admin Consent.
An Intune account with read permissions for authentication during the export.
Internet access to install modules from PSGallery and authenticate against Microsoft Graph.


Module Installation
The script automatically installs all required modules during the first step. No manual installation is needed. Required modules are: PSWriteWord, PSWriteHTML, ImportExcel, Microsoft.Graph.Authentication, Microsoft.Graph.Applications, Microsoft.Graph.Identity.SignIns, AzureAD version 2.0.2.140, and MSAL.PS.
In some environments, running PowerShell as Administrator may be required to install modules in the AllUsers scope.

Project Components
IntuneDocKit.ps1 — The main graphical interface built with Windows Forms. The user enters the Tenant ID, client name, UPN, and selects the execution mode. It also includes an integrated console showing real-time progress and a visual step progress bar.
MasterExport-MEM.ps1 — The master orchestrator. Coordinates the five process steps and handles the logic for each execution mode. Receives parameters from the GUI or directly from the command line.
Export-MEMConfiguration.ps1 — The export engine. Authenticates the user via MSAL Device Code Flow, makes all calls to Microsoft Graph API, and generates the Word document using the included template.
Export-MEMConfiguration.xml — The configuration file. Stores the Tenant ID, the App Registration Client ID, and controls which Intune sections are included in the documentation. This file is updated automatically during Step 4.
MEMDocumentationTempl.docx — The base Word template. Defines the structure, styles, and cover page of the output file.
IntuneDocKit.exe — A compiled executable for easy distribution without needing to run PowerShell directly.

Execution Modes
Mode 1: Create a new App Registration (first run)
Used for the first run on a new tenant. The script creates the App Registration in Azure, assigns all required permissions, grants Admin Consent automatically, updates the XML with the obtained TenantId and AppId, and then proceeds to export the full documentation.
powershell.\MasterExport-MEM.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -ClientName "Client Name"
Mode 2: Use an existing App Registration
For subsequent runs when the App Registration already exists. The script skips the Azure creation and configuration steps, reads the existing ClientId from the XML, and proceeds directly to the export.
powershell.\MasterExport-MEM.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -ClientName "Client Name" -SkipAppCreation
Mode 3: Create App Registration only
Useful when you want to prepare tenant access but run the export at a later time. The script creates the App Registration, assigns permissions, and updates the XML, but does not generate any document.
powershell.\MasterExport-MEM.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -OnlyCreateApp

The Five Process Steps
Step 1 – Modules: Verifies that all required PowerShell modules are installed. Installs any missing ones automatically from PSGallery. Also configures the NuGet provider and sets PSGallery as a trusted repository.
Step 2 – App Registration: Creates an App Registration named "IntuneDocKit" in the Azure tenant using Microsoft.Graph. If an app with that name already exists, it reuses it without creating duplicates. Also enables Public Client Flows and creates the corresponding Service Principal. This step only runs in Modes 1 and 3.
Step 3 – Permissions & Admin Consent: Assigns the required delegated permissions to the App Registration and automatically grants Admin Consent for all tenant users. All assigned permissions are read-only and cover devices, apps, users, groups, directory, policies, and Intune audit areas. This step only runs in Modes 1 and 3.
Step 4 – XML Update: Writes the Tenant ID and Client ID to the Export-MEMConfiguration.xml file so the export can authenticate correctly. In Mode 2, it simply verifies that the values already exist in the XML.
Step 5 – Export: Launches a clean PowerShell session to avoid conflicts between the Microsoft.Graph and MSAL.PS modules. In this session, the user is authenticated via MSAL Device Code Flow, all Microsoft Graph API calls are made using the Beta version, and the Word document with the full tenant documentation is generated.

Device Code Authentication
During Step 5, the script prompts the user to authenticate. The flow is as follows:

The script displays a one-time code in the console.
The user opens a browser and navigates to https://microsoft.com/devicelogin.
The user enters the code shown in the console.
The user authenticates with an Intune account that has read permissions.
Once authentication is complete, the process continues automatically.

The account used in this step does not need Azure permissions — it only needs read access to Intune.

Microsoft Graph Permissions
The following delegated permissions are automatically assigned to the App Registration during Step 3:
openid, profile, offline_access, User.Read, DeviceManagementApps.Read.All, DeviceManagementConfiguration.Read.All, DeviceManagementManagedDevices.Read.All, DeviceManagementRBAC.Read.All, DeviceManagementServiceConfig.Read.All, DeviceManagementScripts.Read.All, Directory.Read.All, User.Read.All, Group.Read.All, Organization.Read.All, Application.Read.All, Policy.Read.All, AuditLog.Read.All, Agreement.Read.All, CloudPC.Read.All, and Policy.ReadWrite.ConditionalAccess.
All permissions are read-only except Policy.ReadWrite.ConditionalAccess, which is required to read Conditional Access policies through the delegated API available in Graph.

XML Configuration File
The Export-MEMConfiguration.xml file controls the export behavior. The most relevant fields are:

Tenant — Stores the client's Tenant ID. Updated automatically in Step 4.
ClientId — Stores the App Registration's App ID. Updated automatically in Step 4.
Document — Enables Word file generation. Default value: True.
DocumentName — Defines the generated file name. Default pattern: IntuneDoc_{ClientName}_{date}.docx.
Export — Enables additional export in JSON format for each section. Default value: False.
ExportPath — Defines the path where exported files are saved. Default: C:\temp\export_{date}.
graphApiVersion — Defines the Graph API version. Default: Beta.

The Process section of the XML contains one entry per Intune section with a value of True or False to include or exclude it from the documentation.

Output
When the process completes, a Word file is generated in the configured path, typically C:\temp. The file opens automatically upon completion. The document includes a cover page with the client name and date, and one section per Intune area configured in the XML, with tables of values and properties extracted directly from the tenant.

GUI Usage
When opening IntuneDocKit.exe or running IntuneDocKit.ps1, a graphical interface appears with the following fields:

Client Name — Appears on the cover page of the generated Word document.
Tenant ID — Identifies the client's Azure tenant.
User UPN — Used as a hint for authentication in Step 5.
Execution Mode — Select from the three modes described above.

The GUI also includes a five-step progress bar, an integrated console showing real-time output with color-coded messages by type, and buttons to copy the equivalent PowerShell command, clear the console, and copy the Device Code during authentication.

Credits
This project is based on the original work by matbe, available at https://github.com/matbe/MEMDocumentAndExporter.
The original project was picked up, modernized, and extended with automation of the App Registration process, permission assignment, Admin Consent, and a complete graphical interface.
