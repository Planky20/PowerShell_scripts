This repository contains a collection of PowerShell scripts that I have used throughout my career as a System Administrator. All scripts have been carefully reviewed to remove any sensitive data.

📁 Repository Structure

- **SPO_Scripts**  
  Contains scripts used to automate user identity lifecycle processes. These integrate PowerShell with Azure App Registrations, Windows Server scheduled tasks, and SharePoint lists.  
  Key functionalities include:
  - Automated account creation, modification, and deletion
  - Operations on Active Directory (AD) objects, later synced to Entra ID via CloudSync
  - Secure credential handling via Azure Key Vault
  - Execution on a Domain Controller (hosted on an Azure VM with a system-assigned managed identity)


- **Other Scripts**  
  A collection of miscellaneous scripts used to streamline day-to-day administrative tasks