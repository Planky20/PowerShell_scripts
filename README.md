# 💼 PowerShell Scripts for IT Automation

This repository contains a collection of PowerShell scripts that I have developed and used throughout my career as a **System Administrator**. All scripts have been carefully reviewed to ensure **no sensitive data** is exposed.

---

## 📁 Repository Structure

### 🔹 `SPO_Scripts` – **Identity Lifecycle Automation for SharePoint Online**

This is the **core focus** of the repository. These scripts automate the full **user identity lifecycle** across **SharePoint Online**, **Active Directory**, and **Entra ID** using a modern and secure automation stack.

#### 🔧 Key Technologies Used:
- **Azure App Registrations** – for secure access to the **Microsoft Graph API**
- **Microsoft Graph API** – to perform operations on **Microsoft 365** users and groups
- **Azure Key Vault** – securely handles **credentials and secrets**
- **Windows Task Scheduler** – automates **timed script execution**
- **Azure Virtual Machines** – scripts run on a **Domain Controller** with a **system-assigned managed identity**
- **CloudSync** (*Entra Connect Sync*) – syncs changes from **on-premises AD** to **Entra ID**
- **SharePoint Lists** – act as **dynamic data sources** and **logs** for the automation

#### 🛠️ Main Functionalities:
- Automated **user creation**, **modification**, and **deletion**
- Operations on **Active Directory objects**, synced to **Entra ID**
- Logic driven by **SharePoint Online** lists
- Credential-free execution with **Azure Managed Identity**

---

### 🔹 `Other Scripts` – Supporting tools for system administration

A collection of **supporting scripts** used to streamline day-to-day administrative tasks:

- **Active Directory scripts** – update user attributes (e.g., **job title**, **street address**, **extension attributes**)
- **Azure tools** – check available **VM SKUs**, region compatibility
- **Autopilot** – register devices manually using extracted **hardware IDs**
- **Microsoft 365 utilities** – check **password change history**, convert **email addresses** to **SAMAccountNames**

---

## 🧩 Feel free to explore, adapt, and contribute!
