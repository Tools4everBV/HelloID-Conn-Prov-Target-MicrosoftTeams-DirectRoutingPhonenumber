# HelloID-Conn-Prov-Target-MicrosoftTeams-DirectRoutingPhonenumber
Repository for HelloID Provisioning Target Connector to Microsoft Teams to set the Direct Routing Phonenumber

<a href="https://github.com/Tools4everBV/HelloID-Conn-Prov-Target-MicrosoftTeams-DirectRoutingPhonenumber/network/members"><img src="https://img.shields.io/github/forks/Tools4everBV/HelloID-Conn-Prov-Target-MicrosoftTeams-DirectRoutingPhonenumber" alt="Forks Badge"/></a>
<a href="https://github.com/Tools4everBV/HelloID-Conn-Prov-Target-MicrosoftTeams-DirectRoutingPhonenumber/pulls"><img src="https://img.shields.io/github/issues-pr/Tools4everBV/HelloID-Conn-Prov-Target-MicrosoftTeams-DirectRoutingPhonenumber" alt="Pull Requests Badge"/></a>
<a href="https://github.com/Tools4everBV/HelloID-Conn-Prov-Target-MicrosoftTeams-DirectRoutingPhonenumber/issues"><img src="https://img.shields.io/github/issues/Tools4everBV/HelloID-Conn-Prov-Target-MicrosoftTeams-DirectRoutingPhonenumber" alt="Issues Badge"/></a>
<a href="https://github.com/Tools4everBV/HelloID-Conn-Prov-Target-MicrosoftTeams-DirectRoutingPhonenumber/graphs/contributors"><img alt="GitHub contributors" src="https://img.shields.io/github/contributors/Tools4everBV/HelloID-Conn-Prov-Target-MicrosoftTeams-DirectRoutingPhonenumber?color=2b9348"></a>

| :information_source: Information |
|:---------------------------|
| This repository contains the connector and configuration code only. The implementer is responsible to acquire the connection details such as username, password, certificate, etc. You might even need to sign a contract or agreement with the supplier before implementing this connector. Please contact the client's application manager to coordinate the connector requirements.       |

<p align="center">
  <img src="https://github.com/Tools4everBV/HelloID-Conn-Prov-Target-MicrosoftTeams-DirectRoutingPhonenumber/blob/main/icon.png?raw=true" width="200" heigth="200">
</p>

## Table of Contents
- [HelloID-Conn-Prov-Target-MicrosoftTeams-DirectRoutingPhonenumber](#helloid-conn-prov-target-microsoftteams-directroutingphonenumber)
  - [Table of Contents](#table-of-contents)
  - [Requirements](#requirements)
  - [Introduction](#introduction)
  - [Getting Started](#getting-started)
    - [Installing the Microsoft Teams PowerShell module](#installing-the-microsoft-teams-powershell-module)
    - [Creating the Microsoft Entra ID App Registration](#creating-the-microsoft-entra-id-app-registration)
    - [Application Registration](#application-registration)
    - [Configuring App Permissions](#configuring-app-permissions)
    - [Assign Microsoft Entra ID roles to the application](#assign-microsoft-entra-id-roles-to-the-application)
    - [Authentication and Authorization](#authentication-and-authorization)
    - [Connection settings](#connection-settings)
      - [Remarks](#remarks)
  - [Getting help](#getting-help)
  - [HelloID docs](#helloid-docs)

## Requirements
- Installed and available **Microsoft Teams PowerShell module (at least 5.7.1)**. Please see the [Microsoft documentation](https://learn.microsoft.com/en-us/microsoftteams/teams-powershell-install) for more information. The download [can be found here](https://www.powershellgallery.com/packages/MicrosoftTeams).
- Required to run **On-Premises** since it is not allowed to import a module with the Cloud Agent.
- An **App Registration in Microsoft Entra ID** is required.

## Introduction
For this connector we have the option to set the Direct Routing Phonenumber of a Microsoft Teams User.
> **Note:** This connector is only for setting the Direct Routing Phonenumber, the Microsoft Teams User has to exist already.

The HelloID connector consists of the template scripts shown in the following table.

| Action     | Action(s) Performed                                         | Comment                                                                                                                                                                  |
| ---------- | ----------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| create.ps1 | Update Direct Routing Phonenumber of a Microsoft Teams User | Optionally you can update other Phonenumber types in Microsoft Teams. However, the connector currently isn't desgined for this and might need additional configuration |
| update.ps1 | Update Direct Routing Phonenumber of a Microsoft Teams User | Optionally you can update other Phonenumber types in Microsoft Teams. However, the connector currently isn't desgined for this and might need additional configuration |

## Getting Started
### Installing the Microsoft Teams PowerShell module
Since we use the cmdlets from the Microsoft Teams PowerShell module, it is required this module is installed and available for the service account.
Please follow the [Microsoft documentation on how to install the module](https://learn.microsoft.com/en-us/microsoftteams/teams-powershell-install). 

### Creating the Microsoft Entra ID App Registration
> _The steps below are based on the [Microsoft documentation](https://learn.microsoft.com/en-us/microsoftteams/teams-powershell-application-authentication) as of the moment of release. The Microsoft documentation should always be leading and is susceptible to change. The steps below might not reflect those changes._

### Application Registration
The first step is to register a new **Microsoft Entra ID Application**. The application is used to connect to Microsoft Teams and to manage permissions.

- Navigate to **App Registrations** in Microsoft Entra ID, and select “New Registration” (**Microsoft Entra Portal > Microsoft Entra ID > App Registration > New Application Registration**).
- Next, give the application a name. In this example we are using “**HelloID - PowerShell - Microsoft Teams - DirectRoutingPhonenumber**” as application name.
- Specify who can use this application (**Accounts in this organizational directory only**).
- Specify the Redirect URI. You can enter any url as a redirect URI value. In this example we used http://localhost because it doesn't have to resolve.
- Click the “**Register**” button to finally create your new application.

Some key items regarding the application are the Application ID (which is the Client ID), the Directory ID (which is the Tenant ID) and Client Secret.

### Configuring App Permissions
The [Microsoft Graph documentation](https://docs.microsoft.com/en-us/graph) provides details on which permission are required for each permission type.

- To assign your application the right permissions, navigate to **Microsoft Entra Portal > Microsoft Entra ID > App Registrations**.
- Select the application we created before, and select “**API Permissions**” or “**View API Permissions**”.
- To assign a new permission to your application, click the “**Add a permission**” button.
- From the “**Request API Permissions**” screen click “**Microsoft Graph**”.
- For this connector the following permissions are used as **Application permissions**:
  - Manage Read organization information ***Organization.Read.All***
- To grant admin consent to our application press the “**Grant admin consent for TENANT**” button.

### Assign Microsoft Entra ID roles to the application
Microsoft Entra ID has more than 50 admin roles available. The **Teams Communications Administrator** role should provide the required permissions to update the Direct Routing Phonenumber of a Microsoft Teams User. However, some actions may not be allowed, such as managing other admin accounts, for this the Global Administrator would be required. Please note that the required role may vary based on your configuration.
- To assign the role(s) to your application, navigate to **Microsoft Entra Portal > Microsoft Entra ID > Roles and administrators**.
- On the Roles and administrators page that opens, find and select one of the supported roles e.g. “**Teams Communications Administrator**” by clicking on the name of the role (not the check box) in the results.
- On the Assignments page that opens, click the “**Add assignments**” button.
- In the Add assignments flyout that opens, **find and select the app that we created before**.
- When you're finished, click **Add**.
- Back on the Assignments page, **verify that the app has been assigned to the role**.

For more information about the permissions, please see the Microsoft docs:
- [Application-based authentication in Teams PowerShell Module](https://learn.microsoft.com/en-us/microsoftteams/teams-powershell-application-authentication#setup-application-based-authentication).
- [Use Microsoft Teams administrator roles to manage Teams](https://learn.microsoft.com/en-us/microsoftteams/using-admin-roles).

### Authentication and Authorization
There are multiple ways to authenticate to the Graph API with each has its own pros and cons, in this example we are using the Authorization Code grant type.

*	First we need to get the **Client ID**, go to the **Microsoft Entra Portal > Microsoft Entra ID > App Registrations**.
*	Select your application and copy the Application (client) ID value.
*	After we have the Client ID we also have to create a **Client Secret**.
*	From the Microsoft Entra ID Portal, go to **Microsoft Entra ID > App Registrations**.
*	Select the application we have created before, and select "**Certificates and Secrets**". 
*	Under “Client Secrets” click on the “**New Client Secret**” button to create a new secret.
*	Provide a logical name for your secret in the Description field, and select the expiration date for your secret.
*	It's IMPORTANT to copy the newly generated client secret, because you cannot see the value anymore after you close the page.
*	At last we need to get the **Tenant ID**. This can be found in the Microsoft Entra ID Portal by going to **Microsoft Entra ID > Overview**.

### Connection settings
The following settings are required to connect to the API.

| Setting   | Description  | Mandatory  |
| --------- | ------------ | ---------- |
| Microsoft Entra ID Tenant ID  | The ID of the Tenant in Microsoft Entra ID    | Yes   |
| Microsoft Entra ID App ID | The ID of the App Registration in Microsoft Entra ID with the Microsoft Teams permissions to the specfied organization    | Yes   |
| Microsoft Entra ID App Secret | The Secret of the App Registration in Microsoft Entra ID with the Microsoft Teams permissions to the specfied organization    | Yes   |
| Only Set Phone Number When Empty  | When toggled, the Phone Number will only be set if there currently is no value set    | No    |
| Toggle debug logging  | When toggled, debug logging will be displayed. Note that this is only meant for debugging, please switch this off when in production | No |

#### Remarks
- Currently we only support setting the Direct Routing Phonenumber of a Microsoft Teams User. Optionally you can update other Phonenumber types in Microsoft Teams. However, the connector currently isn't desgined for this and might need additional configuration.
- Currently we only support the create and update action. Optionally you can also configure the delete action to remove/clear Direct Routing Phonenumber of a Microsoft Teams User. However, the connector currently isn't desgined for this and needs additional configuration.

## Getting help
> _For more information on how to configure a HelloID PowerShell connector, please refer to our [documentation](https://docs.helloid.com/hc/en-us/articles/360012558020-Configure-a-custom-PowerShell-target-system) pages_

> _If you need help, feel free to ask questions on our [forum](https://forum.helloid.com)_

## HelloID docs
The official HelloID documentation can be found at: https://docs.helloid.com/
